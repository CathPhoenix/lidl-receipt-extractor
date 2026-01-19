#!/usr/bin/env python3
"""Lidl receipt images -> bookkeeping rows (Top Bun)

Creates **three rows per receipt** (0.0%, 13.5%, 23.0%) and exports to Excel
matching the column order of a provided Sample.xlsx.

Requirements (Python packages)
------------------------------
pip install pillow pytesseract pandas openpyxl

Also requires Tesseract OCR installed:
- Windows: install Tesseract, then either add it to PATH or pass
  --tesseract-cmd "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
- macOS: brew install tesseract
- Linux: sudo apt-get install tesseract-ocr

Usage
-----
python lidl_receipts_to_excel.py --images "C:\\Receipts\\*.png" "C:\\Receipts\\*.jpg" \
  --out "C:\\Receipts\\lidl_output.xlsx" \
  --sample "C:\\Receipts\\Sample.xlsx" \
  --tesseract-cmd "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

Mapping rules implemented (per user spec)
-----------------------------------------
- Date: read from receipt as DD/MM/YY or DD.MM.YY, written as MM/DD/YY
- Vendor: Lidl
- Invoice Number: image filename (basename)
- TB Reference: blank
- VAT: three rows: 0.0%, 13.5%, 23.0%
- Invoice amount excl. VAT: (Total - VAT.1)
  - For 0.0% row: (Total - VAT.1 - DRS)
- VAT.1: the VAT amount (2nd number) after the VAT line
- DRS: value from "Total Deposits paid" (or "Deposits paid")
  - Only placed on the 0.0% row; otherwise 0
- Total: taken from the VAT summary "Total" column when available; otherwise computed
  - 0.0% row: (Invoice amount excl. VAT) + (VAT.1) + (DRS)
  - 13.5% row: (Invoice amount excl. VAT) + (VAT.1)
  - 23.0% row: (Invoice amount excl. VAT) + (VAT.1)
- Description: Food
- Expense Acc: 5008
- Account Number: Stock: Other Food/Toppings
- Paid by: TopBun - Cash

Notes
-----
- If a VAT rate is missing on a receipt, its taxable/VAT amounts are set to 0.
- This script aims for robust parsing from OCR text, but OCR quality varies.
"""

from __future__ import annotations

import argparse
import glob
import os
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
from PIL import Image, ImageFilter, ImageOps

try:
    import pytesseract
except ImportError as e:
    raise SystemExit("Missing dependency 'pytesseract'. Install with: pip install pytesseract") from e


VAT_RATES: List[float] = [0.0, 13.5, 23.0]

DEFAULT_HEADERS = [
    "Date",
    "Vendor",
    "Invoice Number",
    "TB Reference",
    "VAT",
    "Invoice amount excl. VAT",
    "VAT.1",
    "DRS",
    "Total",
    "Description",
    "Expense Acc",
    "Account Number",
    "Paid by",
    "Comments",
    "Checked",
]


# ---------------- OCR helpers ----------------

def normalise_image_for_ocr(img: Image.Image) -> Image.Image:
    """Convert to grayscale, auto-contrast, and sharpen slightly for better OCR."""
    g = ImageOps.grayscale(img)
    g = ImageOps.autocontrast(g)
    g = g.filter(ImageFilter.SHARPEN)
    return g


def ocr_text(image_path: str, tesseract_cmd: Optional[str] = None) -> str:
    """Run Tesseract OCR on an image and return extracted text."""
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    img = Image.open(image_path)
    img = normalise_image_for_ocr(img)

    # psm 6 = Assume a single uniform block of text.
    text = pytesseract.image_to_string(img, config="--oem 1 --psm 6")
    return text


# ---------------- Parsing helpers ----------------

def _clean_ocr(text: str) -> str:
    """Light cleanup for common OCR issues."""
    # Normalize decimal separators and spacing
    text = text.replace(",", ".")

    # Sometimes OCR drops the dot in 23.0 -> 230; try to reduce that risk is hard.
    # We mainly rely on the presence of '% VAT' as anchor.

    return text


def find_date_mmddyy(text: str) -> str:
    """Find receipt date and return as MM/DD/YY (string). Blank if not found."""
    text = _clean_ocr(text)

    # Typical: "Date: 31/08/25"
    m = re.search(r"Date\s*[:\-]?\s*([0-3]\d[/.][01]\d[/.]\d\d)", text, re.IGNORECASE)
    if m:
        raw = m.group(1)
    else:
        # Fallback: last occurrence of dd.mm.yy or dd/mm/yy
        all_dates = re.findall(r"([0-3]\d[/.][01]\d[/.]\d\d)", text)
        raw = all_dates[-1] if all_dates else ""

    if not raw:
        return ""

    raw = raw.replace(".", "/")
    try:
        dt = datetime.strptime(raw, "%d/%m/%y")
        return dt.strftime("%m/%d/%y")
    except ValueError:
        return ""


# VAT summary line examples in Lidl receipts (OCR varies):
#   A  0.0% VAT   64.63   0.00   64.63
#   C 23.0% VAT    2.48   0.46    2.94
# Some OCR outputs only 2 amounts (often Total then VAT).
# We treat the **2nd amount** as VAT (VAT.1) per user spec, and prefer the
# **3rd amount** as Total when present.
VAT_LINE_RE = re.compile(
    r"^\s*(?:[A-Z]\s+)?(?P<rate>\d{1,2}(?:\.\d)?)\s*%\s*VAT\s+"
    r"(?P<a1>\d+\.\d{2})\s+(?P<a2>\d+\.\d{2})(?:\s+(?P<a3>\d+\.\d{2}))?\s*$",
    re.IGNORECASE | re.MULTILINE,
)


def parse_vat_map(text: str) -> Dict[float, Tuple[float, float]]:
    """Return map: { VAT rate -> (total_amount, vat_amount) }.

    VAT.1 is always taken as the 2nd amount after the VAT rate.
    Total is taken as the 3rd amount when available, otherwise the 1st.
    """
    text = _clean_ocr(text)
    out: Dict[float, Tuple[float, float]] = {}

    for m in VAT_LINE_RE.finditer(text):
        rate_raw = m.group("rate")
        try:
            rate = float(rate_raw)
        except ValueError:
            continue

        a1 = float(m.group("a1"))
        vat = float(m.group("a2"))
        a3 = m.group("a3")
        total = float(a3) if a3 else a1

        # Safety: if OCR swapped columns and 'total' looks smaller than VAT,
        # fall back to the other captured amount.
        if total < vat and a1 >= vat:
            total = a1

        # Normalize expected rates
        if abs(rate - 23.0) < 0.2:
            rate = 23.0
        elif abs(rate - 13.5) < 0.2:
            rate = 13.5
        elif abs(rate - 0.0) < 0.2:
            rate = 0.0

        out[rate] = (total, vat)

    return out


def parse_deposits(text: str) -> float:
    """Return DRS (Total Deposits paid) if present, else 0."""
    text = _clean_ocr(text)

    # Try "Total Deposits paid" first
    m = re.search(r"Total\s+Deposits\s+paid\s*\n?.*?(\d+\.\d{2})", text, re.IGNORECASE | re.DOTALL)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return 0.0

    # Fallback "Deposits paid"
    m2 = re.search(r"Deposits\s+paid\s*\n?.*?(\d+\.\d{2})", text, re.IGNORECASE | re.DOTALL)
    if m2:
        try:
            return float(m2.group(1))
        except ValueError:
            return 0.0

    return 0.0


def format_vat(rate: float) -> str:
    """Format VAT label exactly as user wants."""
    if rate == 13.5:
        return "13.5%"
    # 0.0 and 23.0 keep one decimal
    return f"{rate:.1f}%"


def rows_for_receipt(text: str, image_path: str) -> List[dict]:
    """Build the 3 bookkeeping rows for one receipt OCR text."""
    date_str = find_date_mmddyy(text)
    vat_map = parse_vat_map(text)
    drs = parse_deposits(text)

    image_name = os.path.basename(image_path)

    rows: List[dict] = []
    for rate in VAT_RATES:
        total_from_receipt, vat_amount = vat_map.get(rate, (0.0, 0.0))

        # DRS only on the 0% row
        drs_cell = drs if rate == 0.0 else 0.0

        # Invoice amount excl. VAT is derived from Total minus VAT (and minus DRS for 0%)
        inv_ex_vat = max(total_from_receipt - vat_amount - drs_cell, 0.0)

        # Prefer the receipt's Total when available; otherwise compute.
        if total_from_receipt > 0:
            total = total_from_receipt
        else:
            total = inv_ex_vat + vat_amount + drs_cell

        rows.append(
            {
                "Date": date_str,
                "Vendor": "Lidl",
                "Invoice Number": image_name,
                "TB Reference": "",
                "VAT": format_vat(rate),
                "Invoice amount excl. VAT": round(inv_ex_vat, 2),
                "VAT.1": round(vat_amount, 2),
                "DRS": round(drs_cell, 2),
                "Total": round(total, 2),
                "Description": "Food",
                "Expense Acc": 5008,
                "Account Number": "Stock: Other Food/Toppings",
                "Paid by": "TopBun - Cash",
                "Comments": "",
                "Checked": "x"
            }
        )

    return rows


def load_headers_from_sample(sample_path: Optional[str]) -> Optional[List[str]]:
    """Return column order from Sample.xlsx if provided."""
    if not sample_path:
        return None
    if not os.path.exists(sample_path):
        return None

    try:
        df = pd.read_excel(sample_path, engine="openpyxl")
        cols = list(df.columns)
        return cols if cols else None
    except Exception:
        return None


def expand_image_patterns(patterns: List[str]) -> List[str]:
    files: List[str] = []
    for p in patterns:
        files.extend(glob.glob(p))
    # de-dup preserving order
    out: List[str] = []
    seen = set()
    for f in files:
        if f not in seen and os.path.isfile(f):
            out.append(f)
            seen.add(f)
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description="Extract Lidl receipt totals by VAT band into Excel.")
    ap.add_argument("--images", nargs="+", required=True, help="Image file paths and/or globs (*.png, *.jpg, etc.).")
    ap.add_argument("--out", required=True, help="Output Excel path.")
    ap.add_argument("--sample", default=None, help="Optional Sample.xlsx to copy column order.")
    ap.add_argument("--tesseract-cmd", default=None, help="Path to tesseract binary (Windows), if not on PATH.")

    args = ap.parse_args()

    files = expand_image_patterns(args.images)
    if not files:
        raise SystemExit("No images found for the given --images patterns.")

    all_rows: List[dict] = []
    for fpath in files:
        try:
            text = ocr_text(fpath, tesseract_cmd=args.tesseract_cmd)
        except Exception as e:
            print(f"[WARN] OCR failed for {fpath}: {e}")
            continue

        all_rows.extend(rows_for_receipt(text, fpath))

    if not all_rows:
        raise SystemExit("No rows produced. Check OCR/Tesseract and receipt layout.")

    headers = load_headers_from_sample(args.sample) or DEFAULT_HEADERS
    # Ensure all expected headers exist
    for h in DEFAULT_HEADERS:
        if h not in headers:
            headers.append(h)

    df = pd.DataFrame(all_rows)

    # Ensure any sample-only columns exist (if sample had extra columns)
    for col in headers:
        if col not in df.columns:
            df[col] = ""

    df = df[headers]

    out_path = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    df.to_excel(out_path, index=False, engine="openpyxl")
    print(f"Wrote {len(df)} rows to: {out_path}")


if __name__ == "__main__":
    main()
