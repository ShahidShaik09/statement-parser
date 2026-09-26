"""
Bank Statement PDF -> Excel Parser (v6 - adds OCR support for scanned statements)
====================================================================================

Three extraction methods now:

  "table"       -- real embedded table gridlines (e.g. Axis Bank).
                   Simplest and most reliable when available.
  "line_regex"  -- digital PDF, no real table (e.g. HDFC). Parses
                   words-with-position + regex.
  "ocr_scan"    -- scanned/photographed statement, no text layer at all
                   (e.g. Union Bank via a scanning app). Runs Tesseract
                   OCR, then the same regex + balance-delta approach.

Usage:
    python bank_statement_parser_v6.py Statements.pdf output.xlsx
"""

import sys
import re
import pdfplumber
import pandas as pd

LINE_TOLERANCE = 3

BANK_CONFIGS = {
    "hdfc": {
        "label": "HDFC Bank",
        "method": "line_regex",
        "match_keywords": ["hdfcbank", "hdfc bank"],
        "row_regex": re.compile(
            r"^(?P<date>\d{1,2}/\d{1,2}/\d{2,4})\s+"
            r"(?P<narration>.+?)\s+"
            r"(?P<refno>[A-Za-z0-9]{6,})\s+"
            r"(?P<valuedate>\d{1,2}/\d{1,2}/\d{2,4})\s+"
            r"(?P<amount>[\d,]+\.\d{1,2})\s+"
            r"(?P<balance>[\d,]+\.\d{1,2})\s*$"
        ),
        "footer_markers": [
            "hdfcbank", "closingbalanceincludes", "contentsofthisstatement",
            "registeredofficeaddress", "statementofaccount", "stateaccountbranchgstn",
            "pageno",
        ],
    },

    "axis": {
        "label": "Axis Bank",
        "method": "table",
        "match_keywords": ["axis bank", "axisbank", "ifsc code :utib", "ifsc code:utib"],
        "col_date": 0, "col_narration": 2, "col_debit": 3, "col_credit": 4, "col_balance": 5,
        "date_pattern": re.compile(r"^\d{2}-\d{2}-\d{4}"),
        "opening_balance_markers": ["opening balance"],
        "skip_markers": ["transaction total", "closing balance"],
    },

    "union_bank": {
        "label": "Union Bank of India (scanned/Finacle-style)",
        "method": "ocr_scan",
        "match_keywords": ["union bank of india", "unionbankofindia"],
        "ocr_dpi": 400,
        "ocr_scale": 2.0,
        "ocr_config": "--psm 6",
        "row_regex": re.compile(
            r"^(?P<date>\d{2}\D\d{2}\D\d{4})\S*\s+"
            r"(?P<middle>.+?)\s+"
            r"(?P<amount>\d+(?:[,\s]+\d+)*[.,]\s?\d{2})\s+"
            r"(?P<balance>\d+(?:[,\s]+\d+)*[.,]\s?\d{2})\s*(?P<suffix>CR|DR)\s*$"
        ),
    },

    "default": {
        "label": "Unknown bank (using generic HDFC-style pattern)",
        "method": "line_regex",
        "match_keywords": [],
        "row_regex": re.compile(
            r"^(?P<date>\d{1,2}/\d{1,2}/\d{2,4})\s+"
            r"(?P<narration>.+?)\s+"
            r"(?P<refno>[A-Za-z0-9]{6,})\s+"
            r"(?P<valuedate>\d{1,2}/\d{1,2}/\d{2,4})\s+"
            r"(?P<amount>[\d,]+\.\d{1,2})\s+"
            r"(?P<balance>[\d,]+\.\d{1,2})\s*$"
        ),
        "footer_markers": [],
    },
}


def clean_number(s):
    if not s or not str(s).strip():
        return None
    cleaned = re.sub(r"[^\d.\-]", "", str(s))
    if cleaned in ("", "-", "."):
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def detect_bank(first_page_text):
    text_lower = first_page_text.lower()
    for key, cfg in BANK_CONFIGS.items():
        if key == "default":
            continue
        if any(kw in text_lower for kw in cfg["match_keywords"]):
            return key
    return "default"


def has_text_layer(pdf):
    return bool((pdf.pages[0].extract_text() or "").strip())


def cluster_lines(words):
    words_sorted = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines, current, current_top = [], [], None
    for w in words_sorted:
        if current_top is None or abs(w["top"] - current_top) <= LINE_TOLERANCE:
            current.append(w)
            if current_top is None:
                current_top = w["top"]
        else:
            lines.append(current)
            current, current_top = [w], w["top"]
    if current:
        lines.append(current)
    return lines


def line_text(line_words):
    return " ".join(w["text"] for w in sorted(line_words, key=lambda w: w["x0"]))


def is_footer(line, footer_markers):
    l = line.lower().replace(" ", "")
    return line.strip().startswith("*") or any(m in l for m in footer_markers)


def extract_line_regex(pdf, cfg):
    transactions, current, prev_balance = [], None, None
    for page in pdf.pages:
        words = page.extract_words()
        if not words:
            continue
        for line in cluster_lines(words):
            text = line_text(line).strip()
            m = cfg["row_regex"].match(text)
            if m:
                if current:
                    transactions.append(current)
                amount, balance = clean_number(m.group("amount")), clean_number(m.group("balance"))
                current = {"date": m.group("date"), "narration": m.group("narration"),
                           "debit": None, "credit": None, "balance": balance, "note": ""}
                if prev_balance is None:
                    current["note"] = "FIRST ROW - verify debit/credit manually"
                elif amount is not None and balance is not None:
                    delta = balance - prev_balance
                    if abs(delta - amount) < 0.01:
                        current["credit"] = amount
                    elif abs(delta + amount) < 0.01:
                        current["debit"] = amount
                    else:
                        current["note"] = f"MISMATCH: balance changed by {delta:.2f}, expected +/-{amount:.2f}"
                prev_balance = balance if balance is not None else prev_balance
            else:
                if current is not None and text and not is_footer(text, cfg["footer_markers"]):
                    current["narration"] += " " + text
    if current:
        transactions.append(current)
    return transactions


def extract_table(pdf, cfg):
    transactions, opening_balance, prev_balance = [], None, None
    for page in pdf.pages:
        for table in page.extract_tables():
            for row in table:
                if not row or len(row) <= max(cfg["col_date"], cfg["col_narration"],
                                                cfg["col_debit"], cfg["col_credit"], cfg["col_balance"]):
                    continue
                narration_cell = (row[cfg["col_narration"]] or "").strip()
                narration_lower = narration_cell.lower()
                if any(m in narration_lower for m in cfg["opening_balance_markers"]):
                    opening_balance = clean_number(row[cfg["col_balance"]])
                    prev_balance = opening_balance
                    continue
                if any(m in narration_lower for m in cfg["skip_markers"]):
                    continue
                date_cell = (row[cfg["col_date"]] or "").strip()
                if not cfg["date_pattern"].match(date_cell):
                    continue
                debit, credit, balance = (clean_number(row[cfg["col_debit"]]),
                                           clean_number(row[cfg["col_credit"]]),
                                           clean_number(row[cfg["col_balance"]]))
                note = ""
                if prev_balance is not None and balance is not None:
                    expected = prev_balance + (credit or 0) - (debit or 0)
                    if abs(expected - balance) > 0.02:
                        note = f"MISMATCH: expected balance {expected:.2f}, got {balance:.2f}"
                elif prev_balance is None:
                    note = "No opening balance found - verify manually"
                transactions.append({"date": date_cell, "narration": narration_cell.replace("\n", " "),
                                      "debit": debit, "credit": credit, "balance": balance, "note": note})
                prev_balance = balance if balance is not None else prev_balance
    return transactions


def extract_ocr_scan(pdf_path, cfg):
    from pdf2image import convert_from_path
    import pytesseract
    from pytesseract import Output
    import cv2
    import numpy as np

    def get_lines_native(img):
        data = pytesseract.image_to_data(img, config=cfg["ocr_config"], output_type=Output.DICT)
        lines = {}
        for i in range(len(data['text'])):
            t = data['text'][i].strip()
            if not t:
                continue
            key = (data['block_num'][i], data['par_num'][i], data['line_num'][i])
            lines.setdefault(key, []).append((data['left'][i], t))
        return [" ".join(w[1] for w in sorted(lines[k], key=lambda x: x[0]))
                for k in sorted(lines.keys())]

    def clean_ocr_number(s):
        digits = re.sub(r"\D", "", s)
        return float(digits[:-2] + "." + digits[-2:]) if len(digits) >= 3 else None

    images = convert_from_path(pdf_path, dpi=cfg["ocr_dpi"])
    transactions, prev_signed = [], None

    for page_img in images:
        img = np.array(page_img.convert("L"))
        img = cv2.resize(img, None, fx=cfg["ocr_scale"], fy=cfg["ocr_scale"], interpolation=cv2.INTER_CUBIC)
        for line in get_lines_native(img):
            m = cfg["row_regex"].match(line.strip())
            if not m:
                continue
            amount, balance = clean_ocr_number(m.group("amount")), clean_ocr_number(m.group("balance"))
            suffix = m.group("suffix")
            if amount is None or balance is None:
                continue
            signed = balance if suffix == "CR" else -balance
            note = ""
            debit = credit = None
            if prev_signed is None:
                note = "FIRST ROW - verify manually"
            else:
                delta = signed - prev_signed
                if abs(delta - amount) < 0.02:
                    credit = amount
                elif abs(delta + amount) < 0.02:
                    debit = amount
                else:
                    note = f"MISMATCH (possible OCR misread): balance changed by {delta:.2f}, expected +/-{amount:.2f}"
            transactions.append({"date": m.group("date"), "narration": m.group("middle"),
                                  "debit": debit, "credit": credit, "balance": balance, "note": note})
            prev_signed = signed
    return transactions


def extract(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        if not has_text_layer(pdf):
            print("No text layer detected -- treating as a scanned document (OCR).")
            cfg = BANK_CONFIGS["union_bank"]
            print(f"Using OCR config: {cfg['label']}")
            return extract_ocr_scan(pdf_path, cfg)

        first_text = pdf.pages[0].extract_text() or ""
        bank_key = detect_bank(first_text)
        cfg = BANK_CONFIGS[bank_key]
        print(f"Detected bank: {cfg['label']} (method: {cfg['method']})")
        if cfg["method"] == "table":
            return extract_table(pdf, cfg)
        else:
            return extract_line_regex(pdf, cfg)


def main():
    if len(sys.argv) < 2:
        print("Usage: python bank_statement_parser_v6.py <statement.pdf> [output.xlsx]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "parsed_statement.xlsx"

    print(f"Reading: {pdf_path}")
    transactions = extract(pdf_path)
    print(f"Extracted {len(transactions)} transactions")

    if not transactions:
        print("0 transactions -- this bank's format doesn't match any config yet.")
        return

    flagged = [t for t in transactions if t["note"]]
    print(f"Flagged for review: {len(flagged)} row(s) ({100*len(flagged)/len(transactions):.1f}%)")

    df = pd.DataFrame(transactions)
    df.insert(0, "row", range(1, len(df) + 1))
    df_out = df[["row", "date", "narration", "debit", "credit", "balance", "note"]]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="Transactions", index=False)
        if flagged:
            df_out[df_out["note"] != ""].to_excel(writer, sheet_name="Flagged for review", index=False)

    print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()
