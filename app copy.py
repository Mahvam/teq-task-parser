import os
import re
import uuid
import json
import csv
import html
from tempfile import gettempdir
from pathlib import Path

from flask import Flask, request, render_template, redirect, url_for, send_file

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
        except Exception:
            texts = []
    # If PyPDF2 gave no text, try pdfplumber which can be more robust
    if (not texts) and pdfplumber is not None:
        try:
            with pdfplumber.open(path) as pdf:
                for p in pdf.pages:
                    txt = p.extract_text()
                    if txt:
                        texts.append(txt)
        except Exception:
            pass
    if not texts:
        raise RuntimeError("Unable to extract text from PDF. Ensure the file is a readable text PDF.")
    return "\n".join(texts)


def parse_text_fallback(text):
    """
    Parse task list from PDF text when extraction returns a single line.
    Position names appear before the word "Responsibilities:" and tasks start with •.
    """
    raw = " ".join(text.split())
    results = []

    # Find candidate positions by splitting around "Responsibilities:"
    responsibility_matches = list(re.finditer(r'Responsibilities:', raw, flags=re.IGNORECASE))
    positions = []
    prev_end = 0
    for match in responsibility_matches:
        position_text = raw[prev_end:match.start()].strip()
        if position_text:
            positions.append((match.start(), position_text))
        prev_end = match.end()

    # Extract bullet-based tasks
    for task_match in re.finditer(r'•\s*([^•]+?)(?=(?:•|$))', raw):
        task_text = task_match.group(1).strip()
        importance = 'neutral'

        m_dash = re.search(r'\s*-\s*(critical|high|neutral|low|not\s+important)\s*$', task_text, re.IGNORECASE)
        if m_dash:
            importance = m_dash.group(1).lower()
            task_text = task_text[:m_dash.start()].strip()
        else:
            m_paren = re.search(r'\s*\((critical|high|neutral|low|not\s+important)\)\s*$', task_text, re.IGNORECASE)
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
    try:
        from anthropic import Client
        api_key = os.environ.get('ANTHROPIC_API_KEY')
        if not api_key:
            return None
        client = Client(api_key=api_key)
        prompt = (
            "You are given the full text of a document. Extract all job positions and for each position list its tasks and the importance level for each task. "
            "Return a JSON array of objects with keys: position (string), tasks (array of {task: string, importance: one of critical, high, neutral, low, not important}). "
            "Do NOT include any additional commentary. Here is the document text:\n\n" + text + "\n\nJSON:" 
        )
        resp = client.completions.create(model="claude-2.1", prompt=prompt, max_tokens=2000, temperature=0)
        # response may be in resp.completion or resp['completion'] depending on client version
        completion = None
        if isinstance(resp, dict):
            completion = resp.get('completion') or resp.get('text')
        else:
            completion = getattr(resp, 'completion', None) or getattr(resp, 'text', None) or str(resp)
        if not completion:
            return None
        # attempt to extract JSON substring
        json_text = completion.strip()
        # find first bracket
        idx = json_text.find('[')
        if idx != -1:
            json_text = json_text[idx:]
        try:
            parsed = json.loads(json_text)
            # normalize
            rows = []
            for item in parsed:
                pos = item.get('position') or item.get('title')
                for t in item.get('tasks', []):
                    task_text = t.get('task') if isinstance(t, dict) else str(t)
                    importance = None
                    if isinstance(t, dict):
                        importance = t.get('importance')
                    if not importance:
                        importance = 'neutral'
                    rows.append({'position': pos, 'task': task_text, 'importance': importance})
            return rows
        except Exception:
            return None
    except Exception:
        return None


def parse_text(text):
    # Try anthropic first
    rows = parse_with_anthropic(text)
    if rows:
        return rows
    # fallback
    return parse_text_fallback(text)


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
    path = tmpdir / f.filename
    f.save(path)
    try:
        text = extract_text_from_pdf(str(path))
    except Exception as e:
        return f'Error extracting PDF text: {e}'
    rows = parse_text(text)
    # save exports
    xlsx_path = tmpdir / f'{uid}.xlsx'
    csv_path = tmpdir / f'{uid}.csv'
    write_csv(rows, str(csv_path))
    write_xlsx(rows, str(xlsx_path))
    RESULTS[uid] = {'rows': rows, 'xlsx': str(xlsx_path), 'csv': str(csv_path), 'filename': f.filename}
    return render_template('index.html', results=rows, uid=uid)


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
    path = tmpdir / f.filename
    f.save(path)
    try:
        text = extract_text_from_pdf(str(path))
    except Exception as e:
        return f'Error extracting PDF text: {e}'
    return '<html><body><h1>Raw Extracted Text</h1><pre>' + html.escape(text) + '</pre></body></html>'


@app.route('/download/xlsx/<uid>')
def download_xlsx(uid):
    info = RESULTS.get(uid)
    if not info:
        return 'Not found', 404
    return send_file(info['xlsx'], as_attachment=True, download_name=f'teq_tasks_{uid}.xlsx')


@app.route('/download/csv/<uid>')
def download_csv(uid):
    info = RESULTS.get(uid)
    if not info:
        return 'Not found', 404
    return send_file(info['csv'], as_attachment=True, download_name=f'teq_tasks_{uid}.csv')


if __name__ == '__main__':
    app.run(debug=True)
