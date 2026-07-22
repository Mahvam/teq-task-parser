# TEQ Connect Task Parser

A Python + Flask web app built to eliminate a manual data entry bottleneck at an enterprise mining client.

## The Problem

Every week, Hecla Mining emailed a PDF containing updated job task descriptions. Someone had to manually read through it and type every position, task, and importance level into a spreadsheet by hand — a slow, error-prone process.

## The Solution

I built an internal tool that automates the entire workflow:

1. Upload the PDF through a simple web interface
2. Claude AI (Anthropic API) reads and extracts the structured data
3. Results are displayed in a clean table instantly
4. Download as Excel or CSV with one click

No more manual data entry. No more typos. One click instead of an hour of work.

## Tech Stack

- **Python + Flask** — backend web app
- **Claude AI (Anthropic API)** — intelligent PDF extraction
- **pdfplumber + PyPDF2** — PDF reading
- **openpyxl** — Excel file generation
- **HTML/CSS** — simple browser-based interface

## Project Status

- ✅ Phase 1 — PDF upload, AI extraction, Excel/CSV export (complete)
- 🔄 Phase 2 — Google Sheets integration with duplicate detection on weekly imports (in progress)

## Why I Built This

I'm a Customer Success Manager who identified a recurring operational pain point with a client and built a solution instead of waiting for one. This project reflects my approach to CS: find the friction, eliminate it, and make the process scalable.
