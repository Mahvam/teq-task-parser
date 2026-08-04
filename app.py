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

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None

app = Flask(__name__)

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
log = logging.getLogger('teq-parser')

ANTHROPIC_MODEL = "claude-sonnet-4-6"
IMPORTANCE_WORDS = r'critical|high|neutral|low|not\s+important'
VALID_IMPORTANCE = {'critical', 'high', 'neutral', 'low', 'not important'}
RESULTS = {}

SHEET_ID = '1FKV8SHFGWNslp4zxaQC1mgyYVRRt-dNrCOIcPvVOgys'
CREDENTIALS_FILE = Path(__file__).parent / 'credentials.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def get_sheet():
    if gspread is None:
        log.warning('gspread not installed.')
        return None
    try:
        creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
        if creds_json:
            creds_info = json.loads(creds_json)
            creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        else:
            creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SHEET_ID).sheet1
        log.info('Connected to Google Sheet successfully.')
        return sheet
    except Exception as e:
        log.error('Could not connect to Google Sheet: %s', e)
        return None


def get_existing_tasks(sheet):
    try:
        rows = sheet.get_all_values()
        return {(r[0].strip().lower(), r[1].strip().lower()) for r in rows[1:] if len(r) >= 2}
    except Exception as e:
        log.error('Could not read sheet: %s', e)
        return set()


def append_new_tasks(sheet, new_rows):
    try:
        values = [[r['position'], r['task'], r['importance']] for r in new_rows]
        sheet.append_rows(values, value_input_option='RAW')
        log.info('Appended %d new tasks to Google Sheet.', len(new_rows))
    except Exception as e:
        log.error('Could not append to sheet: %s', e)


def filter_new_rows(rows, existing_tasks):
    return [r for r in rows if (r['position'].strip().lower(), r['task'].strip().lower()) not in existing_tasks]


def extract_text_from_pdf(path):
    texts = []
    if PdfReader is not None:
        try:
            reader = PdfReader(path)
            for p in reader.pages:
                txt = p.extract_text()
                if txt:
                    texts.append(txt)
        except Exception as e:
            log.warning('PyPDF2 failed: %s', e)
            texts = []
    if (not texts) and pdfplumber is not None:
        try:
            with pdfplumber.open(path) as pdf:
                for p in pdf.pages:
                    txt = p.extract_text()
                    if txt:
                        texts.append(txt)
        except Exception as e:
            log.warning('pdfplumber failed: %s', e)
    if not texts:
        raise RuntimeError("Unable to extract text from PDF.")
    return "\n".join(texts)


def clean_position(chunk):
    text = ' '.join(chunk.split())
    if not text:
        return 'Unknown'
    if '•' in text:
        text = text.rsplit('•', 1)[-1].strip()
        marker = re.search(r'(?:-\s*|\()(?:' + IMPORTANCE_WORDS + r')\)?\s*', text, re.IGNORECASE)
        if marker:
            text = text[marker.end():].strip()
        else:
            words = text.split()
            text = ' '.join(words[-8:]) if len(words) > 8 else text
    text = ' '.join(text.split()).strip(' -–—:|')
    return text or 'Unknown'


def parse_text_fallback(text):
    raw = " ".join(text.split())
    results = []
    responsibility_matches = list(re.finditer(r'Responsibilities:', raw, flags=re.IGNORECASE))
    positions = []
    prev_end = 0
    for match in responsibility_matches:
        position_text = clean_position(raw[prev_end:match.start()])
        if position_text:
            positions.append((match.start(), position_text))
        prev_end = match.end()
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
        results.append({'position': position_name, 'task': task_text, 'importance': importance})
    return results


def parse_with_anthropic(text):
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        return None
    try:
        from anthropic import Anthropic
    except ImportError:
        return None
    client = Anthropic(api_key=api_key)
    system_prompt = "You extract structured data from job description documents. You reply with JSON only: no preamble, no explanation, no markdown code fences."
    user_prompt = (
        "Below is the full text of a document listing job positions and their tasks.\n\n"
        "Extract every position, and for each one, every task and that task's importance level.\n\n"
        "Rules:\n"
        "- The position must be the job title ONLY (for example 'Underground Mechanic').\n"
        "- importance must be exactly one of: critical, high, neutral, low, not important.\n"
        "- If a task has no stated importance, use 'neutral'.\n"
        "- Preserve the task wording as written.\n\n"
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
        log.error('Claude API call failed: %s', e)
        return None
    completion = "".join(block.text for block in resp.content if getattr(block, 'type', None) == 'text').strip()
    if not completion:
        return None
    completion = re.sub(r'^```(?:json)?|```$', '', completion.strip(), flags=re.MULTILINE).strip()
    start = completion.find('[')
    end = completion.rfind(']')
    if start == -1 or end == -1:
        return None
    try:
        parsed = json.loads(completion[start:end + 1])
    except json.JSONDecodeError:
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
                importance = 'neutral'
            rows.append({'position': pos, 'task': task_text, 'importance': importance})
    if not rows:
        return None
    log.info('Claude extracted %d tasks across %d positions.', len(rows), len(parsed))
    return rows


def parse_text(text):
    rows = parse_with_anthropic(text)
    if rows:
        return rows, 'claude'
    return parse_text_fallback(text), 'fallback'


def write_xlsx(rows, path):
    if Workbook is None:
        raise RuntimeError('openpyxl is required')
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
    path = tmpdir / f'{uid}_upload.pdf'
    f.save(path)
    try:
        text = extract_text_from_pdf(str(path))
    except Exception as e:
        return f'Error extracting PDF text: {e}'

    rows, method = parse_text(text)

    sheet = get_sheet()
    new_rows = rows
    skipped = 0
    if sheet:
        existing = get_existing_tasks(sheet)
        new_rows = filter_new_rows(rows, existing)
        skipped = len(rows) - len(new_rows)
        if new_rows:
            append_new_tasks(sheet, new_rows)
        log.info('Phase 2: %d new, %d duplicate(s) skipped.', len(new_rows), skipped)

    xlsx_path = tmpdir / f'{uid}.xlsx'
    csv_path = tmpdir / f'{uid}.csv'
    write_csv(new_rows, str(csv_path))
    write_xlsx(new_rows, str(xlsx_path))
    RESULTS[uid] = {'rows': new_rows, 'xlsx': str(xlsx_path), 'csv': str(csv_path), 'filename': f.filename, 'method': method}
    return render_template('index.html', results=new_rows, uid=uid, method=method, skipped=skipped)


def _find_export(uid, extension):
    if not re.fullmatch(r'[0-9a-f]{32}', uid):
        return None
    path = Path(gettempdir()) / 'teq_parser' / f'{uid}.{extension}'
    return path if path.exists() else None


@app.route('/download/xlsx/<uid>')
def download_xlsx(uid):
    path = _find_export(uid, 'xlsx')
    if not path:
        return 'Export not found.', 404
    return send_file(str(path), as_attachment=True, download_name=f'teq_tasks_{uid}.xlsx')


@app.route('/download/csv/<uid>')
def download_csv(uid):
    path = _find_export(uid, 'csv')
    if not path:
        return 'Export not found.', 404
    return send_file(str(path), as_attachment=True, download_name=f'teq_tasks_{uid}.csv')


@app.route('/health')
def health():
    key = os.environ.get('ANTHROPIC_API_KEY')
    if not key:
        return '<h1>No ANTHROPIC_API_KEY found</h1>'
    try:
        from anthropic import Anthropic
        client = Anthropic(api_key=key)
        client.messages.create(model=ANTHROPIC_MODEL, max_tokens=10, messages=[{"role": "user", "content": "Reply with just: ok"}])
        return f'<h1>Claude is reachable</h1>'
    except Exception as e:
        return f'<h1>Claude call failed</h1><pre>{html.escape(str(e))}</pre>'


if __name__ == '__main__':
    app.run(debug=True)
