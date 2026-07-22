# TEQ Connect Task Parser

Simple Flask app to upload a PDF, extract job positions and tasks, and download results.

Usage

1. Install dependencies:

```bash
python -m pip install -r requirements.txt
```

2. (Optional) Set your Anthropic API key to enable Claude parsing:

```bash
export ANTHROPIC_API_KEY=your_key_here
```

3. Run the app:

```bash
python app.py
```

4. Open http://127.0.0.1:5000 in your browser.

If Anthropic isn't configured the app will use a local regex-based fallback parser.
