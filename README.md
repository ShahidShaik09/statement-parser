# Bank Statement PDF to Excel

Converts bank statement PDFs (digital or scanned) into clean Excel files,
auto-detecting the bank and extraction method, with a built-in balance-chain
reconciliation check that flags any row where the math doesn't add up.

## Supported so far
- **HDFC** — digital PDF, line-regex method
- **Axis Bank** — digital PDF, real embedded table
- **Union Bank of India** — scanned statement, OCR (Tesseract), ~78% clean
- Any other bank falls back to a generic pattern and may return 0 transactions
  until a config is added for it in `BANK_CONFIGS`

## Files
- `bank_statement_parser_v6.py` — the core parser (CLI + importable)
- `streamlit_app.py` — web UI wrapping the parser
- `requirements.txt` — Python dependencies
- `packages.txt` — system packages required by Streamlit Community Cloud
  (Tesseract OCR + Poppler, needed for scanned PDFs)

## Run locally

```bash
pip install -r requirements.txt
# Also install system packages (Ubuntu/Debian):
sudo apt-get install tesseract-ocr poppler-utils

streamlit run streamlit_app.py
```

Or use the parser directly from the command line:

```bash
python3 bank_statement_parser_v6.py Statements.pdf output.xlsx
```

## Deploy to Streamlit Community Cloud
1. Push this repo to GitHub (see below).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub.
3. "New app" → select this repo → main file: `streamlit_app.py` → Deploy.
4. `packages.txt` is picked up automatically to install Tesseract/Poppler.

## Adding a new bank
Add an entry to `BANK_CONFIGS` in `bank_statement_parser_v6.py` with the
matching keywords and either a `table`, `line_regex`, or `ocr_scan` method.
See the existing entries for the shape each method expects.
