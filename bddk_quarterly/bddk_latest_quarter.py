#!/usr/bin/env python3
"""
BDDK Quarterly Consumer-Loan Extractor — LATEST quarter, 25 banks
=================================================================

Recurring companion to the historical backfill script. Each run:

  1. Auto-detects the most recent quarter for which any of the 25 tracked
     banks has published a consolidated ("Konsolide") report on BDDK's
     "Bağımsız Denetim Raporları" portal (https://www.bddk.gov.tr/BdrUyg).
  2. Downloads that quarter's consolidated report for each of the 25 banks.
  3. Extracts the "Tüketici kredileri ... ilişkin bilgiler" footnote table
     (10 fields: -TP total + Konut/Taşıt/İhtiyaç/Diğer, -YP total + 4).
  4. Writes a single-quarter workbook in the Template.xlsx layout:
        output/BDDK_25Banks_<QUARTER>.xlsx   (e.g. ..._2026Q2.xlsx)

You never edit this file between quarters — just re-run it. When 2026-Q2 is
published, `python3 bddk_latest_quarter.py` picks it up automatically; same
for 2026-Q3, etc. To pin a specific quarter without editing code:
        python3 bddk_latest_quarter.py --quarter 2026Q2

Extraction robustness (per bank/quarter), in order:
  * MANUAL_OVERRIDES  — a few verified historical cells (kept for exact
    reproduction if ever re-run; inert for future quarters).
  * text / .docx parse — the normal path for almost every bank.
  * OCR fallback      — when a bank renders its table as an image / vector
    curves / letter-spaced text (e.g. Fibabanka), the page is rasterized and
    read with RapidOCR (pip-only, no system deps). OCR numbers are accepted
    ONLY if they pass the internal sum-check
    (Konut+Taşıt+İhtiyaç+Diğer == -TP total, same for -YP); otherwise the
    cell is left 0 and flagged, with the page image saved to _ocr_pages/ for
    a quick manual read.
  * else 0 (bank genuinely has no consumer-loan note, or file missing on
    BDDK's server).

Rules: no consolidated report for the quarter -> 0; blank / "-" field -> 0.

Dependencies (one-time):
    pip install pymupdf rapidocr-onnxruntime pdfplumber openpyxl requests \
                truststore python-docx pillow
No Tesseract / Homebrew required — OCR is pure-pip via RapidOCR.
"""

import argparse
import io
import json
import os
import re
import sys
import time
import urllib.parse
import zipfile
from datetime import datetime

import openpyxl
import pdfplumber
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import docx  # python-docx; a few banks (e.g. TEB) file .docx reports
except ImportError:
    docx = None

try:
    # bddk.gov.tr serves an incomplete certificate chain; truststore routes
    # verification through the OS trust store (certifi's static bundle fails).
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

BASE = "https://www.bddk.gov.tr/BdrUyg"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

FIELD_ORDER = [
    ("TP_Toplam", "Tüketici Kredileri-TP"),
    ("TP_Konut", "Konut Kredisi"),
    ("TP_Tasit", "Taşıt Kredisi"),
    ("TP_Ihtiyac", "İhtiyaç Kredisi"),
    ("TP_Diger", "Diğer"),
    ("YP_Toplam", "Tüketici Kredileri-YP"),
    ("YP_Konut", "Konut Kredisi"),
    ("YP_Tasit", "Taşıt Kredisi"),
    ("YP_Ihtiyac", "İhtiyaç Kredisi"),
    ("YP_Diger", "Diğer"),
]
ZERO_RESULT = {key: 0 for key, _ in FIELD_ORDER}

# The 25 banks to track, in the requested output order. Names must match the
# BankaAdi values returned by KurulusListesiGetir exactly.
WANTED_BANKS = [
    "AKBANK T.A.Ş.",
    "ALBARAKA TÜRK KATILIM BANKASI A.Ş.",
    "ALTERNATİFBANK A.Ş.",
    "ANADOLUBANK A.Ş.",
    "BURGAN BANK A.Ş.",
    "DENİZBANK A.Ş.",
    "FİBABANKA A.Ş.",
    "HAYAT FİNANS KATILIM BANKASI A.Ş.",
    "HSBC BANK A.Ş.",
    "ICBC TURKEY BANK A.Ş.",
    "ING BANK A.Ş.",
    "KUVEYT TÜRK KATILIM BANKASI A.Ş.",
    "QNB BANK A.Ş.",
    "T.C. ZİRAAT BANKASI A.Ş.",
    "TÜRK EKONOMİ BANKASI A.Ş.",
    "TÜRKİYE EMLAK KATILIM BANKASI A.Ş.",
    "TÜRKİYE FİNANS KATILIM BANKASI A.Ş.",
    "TÜRKİYE GARANTİ BANKASI A.Ş.",
    "TÜRKİYE HALK BANKASI A.Ş.",
    "TÜRKİYE VAKIFLAR BANKASI T.A.O.",
    "TÜRKİYE İŞ BANKASI A.Ş.",
    "VAKIF KATILIM BANKASI A.Ş.",
    "YAPI VE KREDİ BANKASI A.Ş.",
    "ZİRAAT KATILIM BANKASI A.Ş.",
    "ŞEKERBANK T.A.Ş.",
]

# Verified historical overrides, keyed by (eft_kodu, year, month). These three
# reports render the consumer-loan table as image/curves/letter-spaced text;
# values were read from the rendered pages and sum-checked. Kept so those exact
# quarters reproduce if this tool is ever pointed at them; the OCR fallback is
# the general path for unknown future quarters.
MANUAL_OVERRIDES = {
    (103, 2023, 9): {"TP_Toplam": 11119514, "TP_Konut": 81896, "TP_Tasit": 498,
                     "TP_Ihtiyac": 11037120, "TP_Diger": 0,
                     "YP_Toplam": 0, "YP_Konut": 0, "YP_Tasit": 0, "YP_Ihtiyac": 0, "YP_Diger": 0},
    (103, 2024, 3): {"TP_Toplam": 11806193, "TP_Konut": 67580, "TP_Tasit": 484,
                     "TP_Ihtiyac": 11738129, "TP_Diger": 0,
                     "YP_Toplam": 0, "YP_Konut": 0, "YP_Tasit": 0, "YP_Ihtiyac": 0, "YP_Diger": 0},
    (206, 2026, 3): {"TP_Toplam": 15012676, "TP_Konut": 4589164, "TP_Tasit": 1597563,
                     "TP_Ihtiyac": 8825949, "TP_Diger": 0,
                     "YP_Toplam": 0, "YP_Konut": 0, "YP_Tasit": 0, "YP_Ihtiyac": 0, "YP_Diger": 0},
}


def sums_ok(result):
    """True iff sub-items sum to the total for both -TP and -YP."""
    for cur in ("TP", "YP"):
        parts = sum(result[f"{cur}_{k}"] for k in ("Konut", "Tasit", "Ihtiyac", "Diger"))
        if result[f"{cur}_Toplam"] != parts:
            return False
    return True


def assert_overrides_valid():
    for key, vals in MANUAL_OVERRIDES.items():
        if not sums_ok(vals):
            raise AssertionError(f"Override {key}: sub-items do not sum to total")


# Header regexes. "T\w*ketici" tolerates OCR/text variants of "Tüketici"
# (Tuketici, Tiketici, doubled letters). Sub-labels already tolerate ş/s,
# ı/i, ç/c, ğ/g and zero-width word gaps (\s*). "ketici"/"Kredisi" are
# distinctive enough that loosening carries negligible false-match risk.
_DASH = r"[-–—]"
TP_HEADER_RE = re.compile(r"^T\w*ketici\s*Kredileri\s*" + _DASH + r"\s*TP\b", re.IGNORECASE)
YP_HEADER_RE = re.compile(r"^T\w*ketici\s*Kredileri\s*" + _DASH + r"\s*YP\b", re.IGNORECASE)
KONUT_RE = re.compile(r"^Konut\s*Kredisi\b", re.IGNORECASE)
TASIT_RE = re.compile(r"^Ta[şs][ıi]t\s*Kredisi\b", re.IGNORECASE)
IHTIYAC_RE = re.compile(r"^[İIi]htiya[çc]\s*Kredisi\b", re.IGNORECASE)
DIGER_RE = re.compile(r"^Di[ğg]er\b", re.IGNORECASE)
NUM_TOKEN_RE = re.compile(r"\(?-?\d[\d,.]*\)?|(?<!\S)-(?!\S)")

# Detects the "Tüketici kredileri, bireysel kredi kartları ..." section header
# in a report's text layer, so we only OCR when the note really exists.
SECTION_HDR_RE = re.compile(r"t[üu]ketici\s+kredileri", re.IGNORECASE)
SECTION_HDR2_RE = re.compile(r"bireysel\s+kredi\s+kart", re.IGNORECASE)


def make_session():
    session = requests.Session()
    session.headers.update(HEADERS)
    retry = Retry(total=5, backoff_factor=1.5, status_forcelist=[429, 500, 502, 503, 504])
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.mount("http://", HTTPAdapter(max_retries=retry))
    return session


def normalize_line(line):
    return line.replace("\xa0", " ").strip()


def parse_number(tok):
    tok = tok.strip()
    if tok in ("-", "–", "—", ""):
        return 0
    neg = tok.startswith("(") and tok.endswith(")")
    if neg:
        tok = tok[1:-1]
    if re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", tok):
        tok = tok.replace(",", "")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", tok):
        tok = tok.replace(".", "").replace(",", ".")
    try:
        val = float(tok)
    except ValueError:
        return 0
    val = -val if neg else val
    return int(val) if val == int(val) else val


def line_last_number(line):
    toks = NUM_TOKEN_RE.findall(line)
    return parse_number(toks[-1]) if toks else None


def parse_lines(lines):
    """Extract the 10 template fields from a list of 'label val1 val2 total'
    lines. Works for pdfplumber text, .docx flattened rows, and OCR-
    reconstructed rows alike. Last number on a label's line = Toplam column.
    Returns (result_dict, status)."""
    result = dict(ZERO_RESULT)
    n = len(lines)

    tp_idx = next((i for i, l in enumerate(lines) if TP_HEADER_RE.match(l)), None)
    if tp_idx is None:
        return result, "TP header not found"
    result["TP_Toplam"] = line_last_number(lines[tp_idx]) or 0

    def find_after(start, pattern, window=8):
        for j in range(start + 1, min(start + 1 + window, n)):
            if pattern.match(lines[j]):
                return j
        return None

    idx = tp_idx
    for key, pat in (("TP_Konut", KONUT_RE), ("TP_Tasit", TASIT_RE),
                     ("TP_Ihtiyac", IHTIYAC_RE), ("TP_Diger", DIGER_RE)):
        f = find_after(idx, pat)
        if f is not None:
            result[key] = line_last_number(lines[f]) or 0
            idx = f

    yp_idx = None
    for j in range(idx + 1, min(idx + 1 + 20, n)):
        if YP_HEADER_RE.match(lines[j]):
            yp_idx = j
            break
    if yp_idx is None:
        return result, "YP header not found"
    result["YP_Toplam"] = line_last_number(lines[yp_idx]) or 0
    idx = yp_idx
    for key, pat in (("YP_Konut", KONUT_RE), ("YP_Tasit", TASIT_RE),
                     ("YP_Ihtiyac", IHTIYAC_RE), ("YP_Diger", DIGER_RE)):
        f = find_after(idx, pat)
        if f is not None:
            result[key] = line_last_number(lines[f]) or 0
            idx = f

    return result, "ok"


def _pdf_to_lines(pdf_path):
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines.extend(normalize_line(l) for l in text.split("\n"))
    return lines


def _docx_to_lines(docx_path):
    doc = docx.Document(docx_path)
    lines = []
    for table in doc.tables:
        for row in table.rows:
            cells = [normalize_line(c.text) for c in row.cells]
            if cells and cells[0]:
                lines.append(" ".join(cells))
    return lines


def extract_text_parse(report_path):
    """Normal text/.docx extraction. Returns (result, status)."""
    if report_path.lower().endswith(".docx"):
        if docx is None:
            return dict(ZERO_RESULT), "docx-not-installed"
        lines = _docx_to_lines(report_path)
    else:
        lines = _pdf_to_lines(report_path)
    return parse_lines(lines)


# --------------------------------------------------------------------------
# OCR fallback (RapidOCR — pip-only, no system dependency)
# --------------------------------------------------------------------------
_OCR_ENGINE = None
_OCR_IMPORT_ERROR = None


def _get_ocr():
    global _OCR_ENGINE, _OCR_IMPORT_ERROR
    if _OCR_ENGINE is None and _OCR_IMPORT_ERROR is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            _OCR_ENGINE = RapidOCR()
        except Exception as e:  # noqa: BLE001 - report any import/init failure
            _OCR_IMPORT_ERROR = e
    return _OCR_ENGINE


def _ocr_png_to_lines(png_path):
    """OCR one page image and reconstruct table rows (cluster boxes by y,
    order by x, join into 'label v1 v2 total' strings)."""
    engine = _get_ocr()
    if engine is None:
        raise RuntimeError(f"OCR engine unavailable: {_OCR_IMPORT_ERROR}")
    result, _ = engine(png_path)
    if not result:
        return []
    items = []
    for box, text, _score in result:
        ys = [p[1] for p in box]
        xs = [p[0] for p in box]
        items.append((sum(ys) / 4.0, min(xs), text))
    items.sort()
    rows, cur, last_y = [], [], None
    for y, x, txt in items:
        if last_y is None or abs(y - last_y) < 18:
            cur.append((x, txt))
        else:
            rows.append(sorted(cur))
            cur = [(x, txt)]
        last_y = y
    if cur:
        rows.append(sorted(cur))
    return [normalize_line(" ".join(t for _x, t in r)) for r in rows]


def _candidate_pages(pdf_path):
    """Pages (0-indexed) holding the consumer-loan section table, plus the
    following page (the table can spill over).

    The section *title* — "Tüketici kredileri, bireysel kredi kartları ..." —
    is what identifies the table page. Requiring both the "tüketici kredileri"
    and "bireysel kredi kart" phrases on the SAME page pinpoints it and avoids
    earlier prose pages that merely mention "tüketici kredileri" in accounting
    policies."""
    pages = set()
    try:
        with pdfplumber.open(pdf_path) as pdf:
            npages = len(pdf.pages)
            for i, page in enumerate(pdf.pages):
                t = page.extract_text() or ""
                if SECTION_HDR_RE.search(t) and SECTION_HDR2_RE.search(t):
                    pages.add(i)
                    if i + 1 < npages:
                        pages.add(i + 1)
    except Exception:
        return []
    return sorted(pages)[:4]


def extract_with_ocr(pdf_path, tag, ocr_dir, dpi=300):
    """Render the section page(s) and OCR them. Returns (result, status).
    Accepts OCR numbers only if the sum-check passes and TP total > 0."""
    import fitz  # PyMuPDF

    candidates = _candidate_pages(pdf_path)
    if not candidates:
        return dict(ZERO_RESULT), "ocr-no-section-page"

    os.makedirs(ocr_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    all_lines, saved = [], []
    try:
        for pno in candidates:
            if pno >= len(doc):
                continue
            pix = doc[pno].get_pixmap(dpi=dpi)
            png = os.path.join(ocr_dir, f"{tag}_p{pno}.png")
            pix.save(png)
            saved.append(png)
            all_lines.extend(_ocr_png_to_lines(png))
    finally:
        doc.close()

    result, status = parse_lines(all_lines)
    if status != "ok" and result["TP_Toplam"] == 0:
        return result, "ocr-parse-failed (page images saved for manual read)"
    if result["TP_Toplam"] <= 0:
        return result, "ocr-parse-failed (no TP total; images saved)"
    if not sums_ok(result):
        return result, "ocr-sumcheck-failed (images saved for manual read)"
    return result, "ok_ocr"


def get_banks(session):
    r = session.post(f"{BASE}/Home/KurulusListesiGetir",
                     data={"kurulusTuruId": 1, "kurulusGrubuId": ""})
    r.raise_for_status()
    return {b["BankaAdi"]: b for b in r.json()}


RAPOR_URL_RE = re.compile(r'raporUrl=([^"&]+)')
KONSOLIDE_FILE_RE = re.compile(r"KONSOLIDE-(\d{4})-(\d{2})\.zip$", re.IGNORECASE)


def get_year_reports(session, eft_kodu, year):
    """{month: raporUrl} of KONSOLIDE reports for one bank/year."""
    params = {"KurulusTuru": 1, "EFTKodu": eft_kodu, "RaporTipi": "KONSOLIDE",
              "DonemYil": year, "DonemAy": 0}
    r = session.get(f"{BASE}/Home/SorguSonuc", params=params)
    r.raise_for_status()
    out = {}
    for raw_url in RAPOR_URL_RE.findall(r.text):
        url = raw_url.replace("&amp;", "&")
        m = KONSOLIDE_FILE_RE.search(requests.utils.unquote(url))
        if m:
            out[int(m.group(2))] = url
    return out


def _find_report_file(root_dir):
    candidates = []
    for root, _dirs, files in os.walk(root_dir):
        for fn in files:
            if fn.lower().endswith((".pdf", ".docx")):
                path = os.path.join(root, fn)
                candidates.append((os.path.getsize(path), path))
    return max(candidates)[1] if candidates else None


def _find_ext(root_dir, ext):
    for root, _dirs, files in os.walk(root_dir):
        for fn in files:
            if fn.lower().endswith(ext):
                return os.path.join(root, fn)
    return None


def download_report(session, rapor_url, cache_dir, cache_key):
    """Download+extract the report (cached), return path to the PDF/DOCX.
    Handles raw-PDF-as-zip, zip-within-zip, and multi-file archives."""
    os.makedirs(cache_dir, exist_ok=True)
    rep_dir = os.path.join(cache_dir, cache_key)
    marker = os.path.join(rep_dir, ".report_path")
    if os.path.exists(marker):
        with open(marker) as f:
            p = f.read().strip()
        if p and os.path.exists(p):
            return p

    decoded_url = urllib.parse.unquote(rapor_url)
    r = session.get(f"{BASE}/Home/DosyaIndir", params={"raporUrl": decoded_url}, timeout=120)
    r.raise_for_status()
    os.makedirs(rep_dir, exist_ok=True)

    if r.content[:4] == b"%PDF":
        report_path = os.path.join(rep_dir, "report.pdf")
        with open(report_path, "wb") as f:
            f.write(r.content)
    else:
        with zipfile.ZipFile(io.BytesIO(r.content)) as z:
            z.extractall(rep_dir)
        for _ in range(3):
            report_path = _find_report_file(rep_dir)
            if report_path:
                break
            nested = _find_ext(rep_dir, ".zip")
            if not nested:
                break
            with zipfile.ZipFile(nested) as z:
                z.extractall(rep_dir)
        report_path = _find_report_file(rep_dir)

    if report_path is None:
        raise RuntimeError(f"No PDF/DOCX report found inside archive for {rapor_url}")
    with open(marker, "w") as f:
        f.write(report_path)
    return report_path


# --------------------------------------------------------------------------
# Latest-quarter detection
# --------------------------------------------------------------------------
def month_to_q(month):
    return month // 3


def q_label(year, month):
    return f"{year}Q{month_to_q(month)}"


def parse_quarter_arg(s):
    m = re.fullmatch(r"(\d{4})Q([1-4])", s.strip(), re.IGNORECASE)
    if not m:
        raise SystemExit(f"--quarter must look like 2026Q2 (got {s!r})")
    year, q = int(m.group(1)), int(m.group(2))
    return year, q * 3


def gather_reports(session, banks, years, delay):
    """{eft: {(year, month): url}} for the given years."""
    out = {}
    for name in WANTED_BANKS:
        eft = banks[name]["EFTKodu"]
        d = {}
        for yr in years:
            try:
                for m, u in get_year_reports(session, eft, yr).items():
                    d[(yr, m)] = u
            except Exception as e:  # noqa: BLE001
                print(f"  WARN: report index fetch failed for {name} {yr}: {e}")
            time.sleep(delay)
        out[eft] = d
    return out


# --------------------------------------------------------------------------
# Workbook (single quarter -> one data column), Template.xlsx layout
# --------------------------------------------------------------------------
LABEL_COL = 6   # F
FIRST_Q_COL = 7  # G
HEADER_FONT = openpyxl.styles.Font(name="Helvetica", size=12, bold=True)
QUARTER_FONT = openpyxl.styles.Font(name="Helvetica", size=12, bold=False)
FIELD_FONT = openpyxl.styles.Font(name="Helvetica", size=12, bold=False)
YELLOW_FILL = openpyxl.styles.PatternFill(fill_type="solid", fgColor="FFFFFF00")
VALUE_ALIGN = openpyxl.styles.Alignment(horizontal="left")


def build_workbook(quarters, bank_results, bank_order):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    row = 1
    first_bank = True
    for bank_name in bank_order:
        name_cell = ws.cell(row=row, column=LABEL_COL, value=bank_name)
        name_cell.font = HEADER_FONT
        name_cell.fill = YELLOW_FILL
        if first_bank:
            for j, (y, q3, _lbl) in enumerate(quarters):
                qcell = ws.cell(row=row, column=FIRST_Q_COL + j, value=f"{y} Q{q3 // 3}")
                qcell.font = QUARTER_FONT
                qcell.fill = YELLOW_FILL
            first_bank = False
        row += 1
        for key, field_label in FIELD_ORDER:
            ws.cell(row=row, column=LABEL_COL, value=field_label).font = FIELD_FONT
            for j, (_y, _m, qlabel) in enumerate(quarters):
                val = bank_results.get(bank_name, {}).get(qlabel, ZERO_RESULT)[key]
                vcell = ws.cell(row=row, column=FIRST_Q_COL + j, value=val)
                vcell.font = FIELD_FONT
                vcell.alignment = VALUE_ALIGN
                vcell.number_format = "#,##0"
            row += 1
    ws.column_dimensions[openpyxl.utils.get_column_letter(LABEL_COL)].width = 35.5
    for j in range(len(quarters)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(FIRST_Q_COL + j)].width = 17
    ws.freeze_panes = ws.cell(row=2, column=FIRST_Q_COL).coordinate
    return wb


def extract_report(report_path, eft, year, month, ocr_dir):
    """Full extraction chain for one bank/quarter: override -> text -> OCR."""
    override = MANUAL_OVERRIDES.get((eft, year, month))
    if override is not None:
        return dict(override), "manual_override"

    result, status = extract_text_parse(report_path)
    if status == "ok":
        return result, "ok"

    # Text parse missed the TP block. If the section header is present, the
    # table is there but not text-readable -> try OCR.
    tag = f"{eft}_{year}_{month:02d}"
    try:
        ocr_result, ocr_status = extract_with_ocr(report_path, tag, ocr_dir)
    except Exception as e:  # noqa: BLE001 - OCR engine missing or render error
        return result, f"needs_review: text-parse '{status}', OCR unavailable ({e})"
    if ocr_status == "ok_ocr":
        return ocr_result, "ok_ocr"
    # OCR attempted but not trusted -> keep zeros, surface for manual read.
    return dict(ZERO_RESULT), f"needs_review: {ocr_status}"


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--quarter", default=None,
                    help="Pin a quarter, e.g. 2026Q2. Default: auto-detect newest.")
    ap.add_argument("--out-dir", default="output")
    ap.add_argument("--cache-dir", default="bddk_cache")
    ap.add_argument("--ocr-dir", default="_ocr_pages")
    ap.add_argument("--delay", type=float, default=0.4)
    args = ap.parse_args()

    assert_overrides_valid()
    session = make_session()

    print("Fetching bank list...")
    banks = get_banks(session)
    missing = [n for n in WANTED_BANKS if n not in banks]
    if missing:
        raise SystemExit("WANTED_BANKS not found in bank list: " + "; ".join(missing))

    if args.quarter:
        year, month = parse_quarter_arg(args.quarter)
        print(f"Using pinned quarter {q_label(year, month)}.")
        reports = gather_reports(session, banks, [year], args.delay)
    else:
        cur = datetime.now().year
        print(f"Detecting latest available quarter (scanning {cur-1}-{cur})...")
        reports = gather_reports(session, banks, [cur, cur - 1], args.delay)
        all_ym = set().union(*[set(d) for d in reports.values()]) if reports else set()
        if not all_ym:
            raise SystemExit("No consolidated reports found for any of the 25 banks.")
        year, month = max(all_ym)

    target = (year, month)
    have = sum(1 for eft_map in reports.values() if target in eft_map)
    qlbl = q_label(year, month)
    print(f"\nLatest quarter: {qlbl}  —  {have}/{len(WANTED_BANKS)} banks have a report.")
    if have < len(WANTED_BANKS) / 2:
        print("  WARNING: fewer than half the banks have filed this quarter yet — "
              "filing may still be in progress; consider re-running later or pass "
              "--quarter for an earlier one.")

    quarters = [(year, month, qlbl)]
    os.makedirs(args.out_dir, exist_ok=True)
    log_path = os.path.join(args.out_dir, f"run_{qlbl}.jsonl")
    out_path = os.path.join(args.out_dir, f"BDDK_25Banks_{qlbl}.xlsx")

    bank_results = {}
    summary = {"ok": 0, "ok_ocr": 0, "manual_override": 0,
               "no_report": 0, "needs_review": [], "error": []}

    with open(log_path, "w", encoding="utf-8") as logf:
        for i, name in enumerate(WANTED_BANKS, 1):
            eft = banks[name]["EFTKodu"]
            print(f"[{i}/{len(WANTED_BANKS)}] {name} (EFT {eft})")
            rapor_url = reports[eft].get(target)
            if not rapor_url:
                bank_results[name] = {qlbl: dict(ZERO_RESULT)}
                summary["no_report"] += 1
                logf.write(json.dumps({"bank": name, "eft": eft, "quarter": qlbl,
                                       "status": "no_consolidated_report"}, ensure_ascii=False) + "\n")
                continue
            try:
                cache_key = f"{eft}_{year}_{month:02d}"
                report_path = download_report(session, rapor_url, args.cache_dir, cache_key)
                result, status = extract_report(report_path, eft, year, month, args.ocr_dir)
                bank_results[name] = {qlbl: result}
                logf.write(json.dumps({"bank": name, "eft": eft, "quarter": qlbl,
                                       "status": status}, ensure_ascii=False) + "\n")
                if status == "ok":
                    summary["ok"] += 1
                elif status == "ok_ocr":
                    summary["ok_ocr"] += 1
                    print(f"  OCR: read from page image, sum-check passed.")
                elif status == "manual_override":
                    summary["manual_override"] += 1
                else:
                    summary["needs_review"].append((name, status))
                    print(f"  NEEDS REVIEW {qlbl}: {status}")
            except Exception as e:  # noqa: BLE001
                bank_results[name] = {qlbl: dict(ZERO_RESULT)}
                summary["error"].append((name, str(e)))
                logf.write(json.dumps({"bank": name, "eft": eft, "quarter": qlbl,
                                       "status": f"error: {e}"}, ensure_ascii=False) + "\n")
                print(f"  ERROR {qlbl}: {e} -> 0")
            time.sleep(args.delay)

    build_workbook(quarters, bank_results, WANTED_BANKS).save(out_path)

    print("\n" + "=" * 60)
    print(f"Quarter {qlbl}: wrote {out_path}")
    print(f"  text-parsed: {summary['ok']}   OCR: {summary['ok_ocr']}   "
          f"override: {summary['manual_override']}   no-report(0): {summary['no_report']}")
    if summary["needs_review"]:
        print(f"  NEEDS MANUAL REVIEW ({len(summary['needs_review'])}) — "
              f"page images in {args.ocr_dir}/:")
        for name, st in summary["needs_review"]:
            print(f"    - {name}: {st}")
    if summary["error"]:
        print(f"  ERRORS ({len(summary['error'])}) — left as 0 "
              f"(usually a file missing on BDDK's server):")
        for name, e in summary["error"]:
            print(f"    - {name}: {e}")
    print(f"  log: {log_path}")


if __name__ == "__main__":
    main()
