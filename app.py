import os
import re
import uuid
import json
import csv
import html
import logging
from tempfile import gettempdir
from pathlib import Path

from flask import Flask, request, render_template, redirect, url_for, send_file

# Load ANTHROPIC_API_KEY from a .env file sitting next to this script, if present.
# Without this, a key stored in .env never reaches os.environ and the Claude
# path silently gives up. pip install python-dotenv
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / '.env')
except Exception:
    pass

try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None

try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None

app = Flask(__name__)

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
log = logging.getLogger('teq-parser')

# The model doing the extraction. Verify current model names at:
# https://platform.claude.com/docs/en/about-claude/model-deprecations
ANTHROPIC_MODEL = "claude-sonnet-4-6"

IMPORTANCE_WORDS = r'critical|high|neutral|low|not\s+important'
VALID_IMPORTANCE = {'critical', 'high', 'neutral', 'low', 'not important'}

# In-memory store for results: id -> {rows, xlsx_path, csv_path}
RESULTS = {}


def extract_text_from_pdf(path):
    texts = []
    # First try PyPDF2
    if PdfReader is not None:
        try:
            reader = PdfReader(path)
            for p in reader.pages:
                txt = p.extract_text()
                if txt:
                    texts.append(txt)
        except Exception as e:
            log.warning('PyPDF2 could not read the PDF (%s). Trying pdfplumber.', e)
            texts = []
    # If PyPDF2 gave no text, try pdfplumber which can be more robust
    if (not texts) and pdfplumber is not None:
        try:
            with pdfplumber.open(path) as pdf:
                for p in pdf.pages:
                    txt = p.extract_text()
                    if txt:
                        texts.append(txt)
        except Exception as e:
            log.warning('pdfplumber could not read the PDF either: %s', e)
    if not texts:
        raise RuntimeError("Unable to extract text from PDF. Ensure the file is a readable text PDF.")
    return "\n".join(texts)


def clean_position(chunk):
    """
    Tidy up a position name pulled out by the fallback parser.

    The raw chunk runs from the end of the previous "Responsibilities:" to the
    start of the next one, so it carries the PREVIOUS role's bullet text along
    with the new job title. Everything after the final bullet is the new title,
    give or take the tail end of that last task.

    This is a best-effort guess. The Claude path below does this properly.
    """
    text = ' '.join(chunk.split())
    if not text:
        return 'Unknown'

    if '•' in text:
        # Keep only what follows the last bullet.
        text = text.rsplit('•', 1)[-1].strip()
        # That leftover still starts with the tail of a task, which normally
        # ends at its importance marker. Cut everything up to and including it.
        marker = re.search(
            r'(?:-\s*|\()(?:' + IMPORTANCE_WORDS + r')\)?\s*',
            text,
            re.IGNORECASE,
        )
        if marker:
            text = text[marker.end():].strip()
        else:
            # No marker to cut at, so guess: job titles are short. Take the tail.
            words = text.split()
            text = ' '.join(words[-8:]) if len(words) > 8 else text

    text = ' '.join(text.split()).strip(' -–—:|')
    return text or 'Unknown'


def parse_text_fallback(text):
    """
    Parse task list from PDF text when extraction returns a single line.
    Position names appear before the word "Responsibilities:" and tasks start with •.

    This runs only when the Claude call is unavailable. It is pattern matching,
    not understanding, so expect rough edges on the Position column.
    """
    raw = " ".join(text.split())
    results = []

    # Find candidate positions by splitting around "Responsibilities:"
    responsibility_matches = list(re.finditer(r'Responsibilities:', raw, flags=re.IGNORECASE))
    positions = []
    prev_end = 0
    for match in responsibility_matches:
        position_text = clean_position(raw[prev_end:match.start()])
        if position_text:
            positions.append((match.start(), position_text))
        prev_end = match.end()

    # Extract bullet-based tasks
    for task_match in re.finditer(r'•\s*([^•]+?)(?=(?:•|$))', raw):
        task_text = task_match.group(1).strip()
        importance = 'neutral'

        m_dash = re.search(r'\s*-\s*(' + IMPORTANCE_WORDS + r')\s*$', task_text, re.IGNORECASE)
        if m_dash:
            importance = m_dash.group(1).lower()
            task_text = task_text[:m_dash.start()].strip()
        else:
            m_paren = re.search(r'\s*\((' + IMPORTANCE_WORDS + r')\)\s*$', task_text, re.IGNORECASE)
            if m_paren:
                importance = m_paren.group(1).lower()
                task_text = task_text[:m_paren.start()].strip()

        position_name = 'Unknown'
        for pos_start, pos_text in positions:
            if pos_start < task_match.start():
                position_name = pos_text
            else:
                break

        results.append({
            'position': position_name,
            'task': task_text,
            'importance': importance
        })

    return results


def parse_with_anthropic(text):
    """
    Ask Claude to read the document and pull out positions, tasks, importance.

    Returns a list of rows, or None if Claude could not be reached. Every failure
    here is logged loudly on purpose: a silent failure means the fallback quietly
    takes over and you never find out the AI path is dead.
    """
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        log.error(
            'ANTHROPIC_API_KEY is not set, so Claude cannot be called. '
            'Falling back to pattern matching. Put the key in a .env file next to app.py.'
        )
        return None

    try:
        from anthropic import Anthropic
    except ImportError:
        log.error('The anthropic package is not installed. Run: pip install anthropic')
        return None

    client = Anthropic(api_key=api_key)

    system_prompt = (
        "You extract structured data from job description documents. "
        "You reply with JSON only: no preamble, no explanation, no markdown code fences."
    )

    user_prompt = (
        "Below is the full text of a document listing job positions and their tasks.\n\n"
        "Extract every position, and for each one, every task and that task's importance level.\n\n"
        "Rules:\n"
        "- The position must be the job title ONLY (for example 'Underground Mechanic'), "
        "not the surrounding header text, not the word 'Responsibilities'.\n"
        "- importance must be exactly one of: critical, high, neutral, low, not important.\n"
        "- If a task has no stated importance, use 'neutral'.\n"
        "- Preserve the task wording as written. Do not summarise or reword tasks.\n\n"
        "Return a JSON array shaped like:\n"
        '[{"position": "Job Title", "tasks": [{"task": "task text", "importance": "critical"}]}]\n\n'
        "Document text:\n\n" + text
    )

    try:
        resp = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=8000,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )
    except Exception as e:
        log.error('Claude API call failed (%s: %s). Falling back to pattern matching.',
                  type(e).__name__, e)
        return None

    # Pull the text out of the response content blocks.
    completion = "".join(
        block.text for block in resp.content if getattr(block, 'type', None) == 'text'
    ).strip()

    if not completion:
        log.error('Claude returned an empty response. Falling back to pattern matching.')
        return None

    # Strip code fences if the model added them anyway, then find the JSON array.
    completion = re.sub(r'^```(?:json)?|```$', '', completion.strip(), flags=re.MULTILINE).strip()
    start = completion.find('[')
    end = completion.rfind(']')
    if start == -1 or end == -1:
        log.error('No JSON array found in Claude response. Falling back to pattern matching.')
        return None

    try:
        parsed = json.loads(completion[start:end + 1])
    except json.JSONDecodeError as e:
        log.error('Could not parse Claude response as JSON (%s). Falling back.', e)
        return None

    rows = []
    for item in parsed:
        if not isinstance(item, dict):
            continue
        pos = (item.get('position') or item.get('title') or 'Unknown').strip()
        for t in item.get('tasks', []):
            if isinstance(t, dict):
                task_text = (t.get('task') or '').strip()
                importance = (t.get('importance') or 'neutral').strip().lower()
            else:
                task_text = str(t).strip()
                importance = 'neutral'
            if not task_text:
                continue
            if importance not in VALID_IMPORTANCE:
                log.warning('Unexpected importance value %r, using neutral.', importance)
                importance = 'neutral'
            rows.append({'position': pos, 'task': task_text, 'importance': importance})

    if not rows:
        log.error('Claude responded but no tasks came out of it. Falling back.')
        return None

    log.info('Claude extracted %d tasks across %d positions.', len(rows), len(parsed))
    return rows


def parse_text(text):
    rows = parse_with_anthropic(text)
    if rows:
        return rows, 'claude'
    log.warning('Using the pattern-matching fallback. Position names may be messy.')
    return parse_text_fallback(text), 'fallback'


def write_xlsx(rows, path):
    if Workbook is None:
        raise RuntimeError('openpyxl is required to write xlsx files')
    wb = Workbook()
    ws = wb.active
    ws.append(['Position', 'Task', 'Importance Level'])
    for r in rows:
        ws.append([r['position'], r['task'], r['importance']])
    wb.save(path)


def write_csv(rows, path):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Position', 'Task', 'Importance Level'])
        for r in rows:
            writer.writerow([r['position'], r['task'], r['importance']])


@app.route('/')
def index():
    return redirect(url_for('upload'))


@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'GET':
        return render_template('index.html')
    f = request.files.get('file')
    if not f:
        return render_template('index.html')
    uid = uuid.uuid4().hex
    tmpdir = Path(gettempdir()) / 'teq_parser'
    tmpdir.mkdir(parents=True, exist_ok=True)
    # Use the generated id for the saved file so a strange filename cannot
    # write outside the temp folder.
    path = tmpdir / f'{uid}_upload.pdf'
    f.save(path)
    try:
        text = extract_text_from_pdf(str(path))
    except Exception as e:
        return f'Error extracting PDF text: {e}'
    rows, method = parse_text(text)
    # save exports
    xlsx_path = tmpdir / f'{uid}.xlsx'
    csv_path = tmpdir / f'{uid}.csv'
    write_csv(rows, str(csv_path))
    write_xlsx(rows, str(xlsx_path))
    RESULTS[uid] = {
        'rows': rows,
        'xlsx': str(xlsx_path),
        'csv': str(csv_path),
        'filename': f.filename,
        'method': method,
    }
    return render_template('index.html', results=rows, uid=uid, method=method)


@app.route('/debug', methods=['GET', 'POST'])
def debug():
    if request.method == 'GET':
        return '''
            <html><body>
            <h1>PDF Debug Extract</h1>
            <form method="post" enctype="multipart/form-data">
              <input type="file" name="file" accept="application/pdf" required>
              <button type="submit">Upload PDF</button>
            </form>
            </body></html>
        '''
    f = request.files.get('file')
    if not f:
        return 'No file uploaded', 400
    tmpdir = Path(gettempdir()) / 'teq_parser'
    tmpdir.mkdir(parents=True, exist_ok=True)
    path = tmpdir / f'{uuid.uuid4().hex}_debug.pdf'
    f.save(path)
    try:
        text = extract_text_from_pdf(str(path))
    except Exception as e:
        return f'Error extracting PDF text: {e}'
    return '<html><body><h1>Raw Extracted Text</h1><pre>' + html.escape(text) + '</pre></body></html>'


def _find_export(uid, extension):
    """
    Locate an exported file on disk.

    The original code looked this up in RESULTS, which only lives in memory and
    is wiped every time the app restarts. The file itself sits in the temp folder
    named after the uid, so look there instead. The disk is the truth.
    """
    # uid comes straight from the URL, so make sure it cannot wander elsewhere.
    if not re.fullmatch(r'[0-9a-f]{32}', uid):
        return None
    path = Path(gettempdir()) / 'teq_parser' / f'{uid}.{extension}'
    return path if path.exists() else None


@app.route('/download/xlsx/<uid>')
def download_xlsx(uid):
    path = _find_export(uid, 'xlsx')
    if not path:
        return 'That export is no longer on disk. Re-upload the PDF to rebuild it.', 404
    return send_file(str(path), as_attachment=True, download_name=f'teq_tasks_{uid}.xlsx')


@app.route('/download/csv/<uid>')
def download_csv(uid):
    path = _find_export(uid, 'csv')
    if not path:
        return 'That export is no longer on disk. Re-upload the PDF to rebuild it.', 404
    return send_file(str(path), as_attachment=True, download_name=f'teq_tasks_{uid}.csv')


@app.route('/health')
def health():
    """Quick check that the Claude side of things is wired up correctly."""
    key = os.environ.get('ANTHROPIC_API_KEY')
    if not key:
        return '<h1>❌ No ANTHROPIC_API_KEY found</h1><p>Claude cannot be called. The parser will fall back to pattern matching.</p>'
    try:
        from anthropic import Anthropic
        client = Anthropic(api_key=key)
        client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=10,
            messages=[{"role": "user", "content": "Reply with just: ok"}],
        )
        return f'<h1>✅ Claude is reachable</h1><p>Key found, model <code>{ANTHROPIC_MODEL}</code> responded.</p>'
    except Exception as e:
        return f'<h1>❌ Claude call failed</h1><pre>{html.escape(type(e).__name__)}: {html.escape(str(e))}</pre>'


if __name__ == '__main__':
    app.run(debug=True)
