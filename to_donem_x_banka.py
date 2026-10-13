#!/usr/bin/env python3
"""
Dönem x Banka — long-format converter
=====================================

Converts the per-bank workbooks produced by `bddk_consolidated_loans.py` and
`bddk_quarterly/bddk_latest_quarter.py` (one block of 10 rows per bank, one
column per quarter, two sheets) into the single-sheet **"Dönem x Banka"**
layout:

    Dönem   | Kredi Tipi | Kategori                         | <bank> ... | TOPLAM
    2023 Q1 | Tüketici   | Tüketici Kredileri (TP+YP)       | 116540245  | =SUM(...)
    2023 Q1 | Tüketici   | Konut Kredisi                    |  24522089  | =SUM(...)
    ...

Each source figure is the **TP + YP sum** — the source keeps the two currency
blocks apart, this layout adds them. Ten rows per quarter: a Tüketici block
(total + Konut/Taşıt/İhtiyaç/Diğer) and a Ticari block (total +
İşyeri/Taşıt/İhtiyaç/Diğer).

Usage
-----
Rebuild the whole history from the historical backfill workbook:

    python3 to_donem_x_banka.py \
        --source BDDK_Konsolide_25Banks_2023Q1-2026Q1.xlsx \
        --out Tuketici_Ticari_Krediler_Donem_x_Banka.xlsx

Append a new quarter to an existing Dönem x Banka workbook, keeping its
column set and ordering:

    python3 to_donem_x_banka.py \
        --source bddk_quarterly/output/BDDK_25Banks_2026Q2.xlsx \
        --into Tuketici_Ticari_Krediler_Donem_x_Banka.xlsx \
        --out Tuketici_Ticari_Krediler_Donem_x_Banka.xlsx

Add a bank column that the existing workbook does not have yet (it is placed
in the canonical bank order, and back-filled with 0 for earlier quarters, per
the project rule "no consolidated report -> 0"):

    python3 to_donem_x_banka.py --source ... --into ... --out ... \
        --add-bank "ENPARA BANK A.Ş."

By default the sheet carries the 18 banks of the existing report (the 25
tracked banks minus the 7 katılım banks). `--banks` overrides that for a
fresh build.
"""

import argparse
import re
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

SHEET = "Dönem x Banka"

# Source sheet -> (Kredi Tipi, the five Kategori labels in order).
# The first label is the TP+YP total row, the rest are its sub-items.
BLOCKS = [
    ("Tüketici Kredileri", "Tüketici",
     ["Tüketici Kredileri (TP+YP)", "Konut Kredisi", "Taşıt Kredisi",
      "İhtiyaç Kredisi", "Diğer"]),
    ("Taksitli Ticari Krediler", "Ticari",
     ["Taksitli Ticari Krediler (TP+YP)", "İşyeri Kredisi", "Taşıt Kredisi",
      "İhtiyaç Kredisi", "Diğer"]),
]

# Canonical bank ordering (Turkish collation) — the order both extractor
# scripts emit. A bank added with --add-bank is slotted in at its place here.
CANONICAL_BANKS = [
    "AKBANK T.A.Ş.",
    "ALBARAKA TÜRK KATILIM BANKASI A.Ş.",
    "ALTERNATİFBANK A.Ş.",
    "ANADOLUBANK A.Ş.",
    "BURGAN BANK A.Ş.",
    "DENİZBANK A.Ş.",
    "ENPARA BANK A.Ş.",
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

# The 18 banks the published report covers: the tracked 25 less the katılım
# banks, which the report keeps out of this comparison.
KATILIM = [b for b in CANONICAL_BANKS if "KATILIM" in b or b.startswith("HAYAT")]
DEFAULT_BANKS = [b for b in CANONICAL_BANKS
                 if b not in KATILIM and b != "ENPARA BANK A.Ş."]

HEADER_FILL = PatternFill(fill_type="solid", fgColor="FF1F3864")
HEADER_FONT = Font(name="Aptos Narrow", size=11, bold=True, color="FFFFFFFF")
TOTAL_FILL = PatternFill(fill_type="solid", fgColor="FFDDEBF7")
TOTAL_FONT = Font(name="Aptos Narrow", size=11, bold=True)
BODY_FONT = Font(name="Aptos Narrow", size=11)
NUMFMT = "#,##0"

QUARTER_RE = re.compile(r"^(\d{4})\s*Q([1-4])$")


def quarter_key(label):
    m = QUARTER_RE.match(str(label).strip())
    if not m:
        raise ValueError(f"not a quarter label: {label!r}")
    return int(m.group(1)), int(m.group(2))


# --------------------------------------------------------------------------
# Reading the per-bank source workbooks
# --------------------------------------------------------------------------
def read_source(path):
    """-> ({(quarter, bank, tipi, kategori): value}, [quarters], [banks])

    The source layout is: column F holds labels, columns G.. hold quarters
    (their headers sit on row 1). A row whose label matches no field label
    starts a new bank block of 10 rows (TP total + 4 subs, YP total + 4 subs).
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    data, quarters, banks = {}, [], []

    for sheet, tipi, cats in BLOCKS:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]

        # Quarter headers: row 1, from column G rightwards.
        cols = []
        for c in range(7, ws.max_column + 1):
            v = ws.cell(1, c).value
            if v and QUARTER_RE.match(str(v).strip()):
                q = str(v).strip()
                cols.append((c, q))
                if q not in quarters:
                    quarters.append(q)
        if not cols:
            raise SystemExit(f"{path}: no quarter headers on row 1 of {sheet!r}")

        # The 10 field labels of this table, in the order the source writes
        # them: TP total, 4 subs, YP total, the same 4 subs.
        tp_total, subs = cats[0], cats[1:]
        src_tp = tp_total.replace(" (TP+YP)", "-TP")
        src_yp = tp_total.replace(" (TP+YP)", "-YP")
        field_labels = [src_tp] + subs + [src_yp] + subs

        bank = None
        pos = 0
        for r in range(1, ws.max_row + 1):
            label = ws.cell(r, 6).value
            if label is None:
                continue
            label = str(label).strip()
            if pos < len(field_labels) and label == field_labels[pos] and bank:
                # Sub-item labels repeat across the TP and YP halves, so the
                # TP+YP sum is accumulated rather than assigned.
                cat = cats[0] if pos in (0, 5) else label
                for c, q in cols:
                    v = ws.cell(r, c).value
                    key = (q, bank, tipi, cat)
                    data[key] = (data.get(key) or 0) + (v or 0)
                pos += 1
            else:
                bank = label
                pos = 0
                if bank not in banks:
                    banks.append(bank)

    quarters.sort(key=quarter_key)
    return data, quarters, banks


def read_existing(path):
    """-> ({(quarter, bank, tipi, kategori): value}, [quarters], [bank columns])"""
    wb = openpyxl.load_workbook(path, data_only=True)
    if SHEET not in wb.sheetnames:
        raise SystemExit(f"{path}: no {SHEET!r} sheet")
    ws = wb[SHEET]

    banks = []
    for c in range(4, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v is None or str(v).strip().upper() == "TOPLAM":
            continue
        banks.append(str(v).strip())

    data, quarters = {}, []
    for r in range(2, ws.max_row + 1):
        q = ws.cell(r, 1).value
        if not q:
            continue
        q, tipi, cat = str(q).strip(), ws.cell(r, 2).value, ws.cell(r, 3).value
        if q not in quarters:
            quarters.append(q)
        for i, bank in enumerate(banks):
            data[(q, bank, tipi, cat)] = ws.cell(r, 4 + i).value or 0

    quarters.sort(key=quarter_key)
    return data, quarters, banks


# --------------------------------------------------------------------------
# Writing
# --------------------------------------------------------------------------
def build(data, quarters, banks):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET

    headers = ["Dönem", "Kredi Tipi", "Kategori"] + banks + ["TOPLAM"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    last = openpyxl.utils.get_column_letter(3 + len(banks))
    r = 2
    for q in quarters:
        for _sheet, tipi, cats in BLOCKS:
            for cat in cats:
                is_total = cat is cats[0]
                ws.cell(r, 1, q)
                ws.cell(r, 2, tipi)
                ws.cell(r, 3, cat)
                for i, bank in enumerate(banks):
                    ws.cell(r, 4 + i, data.get((q, bank, tipi, cat), 0) or 0)
                ws.cell(r, 4 + len(banks), f"=SUM(D{r}:{last}{r})")
                for c in range(1, 5 + len(banks)):
                    cell = ws.cell(r, c)
                    cell.font = TOTAL_FONT if is_total else BODY_FONT
                    if is_total:
                        cell.fill = TOTAL_FILL
                    if c >= 4:
                        cell.number_format = NUMFMT
                r += 1

    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(4 + len(banks))}{r - 1}"
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 11
    ws.column_dimensions["C"].width = 32
    # One span for the bank columns plus TOPLAM, as Excel writes it.
    span = ws.column_dimensions["D"]
    span.width = 15
    span.max = 4 + len(banks)
    return wb


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--source", required=True,
                    help="Per-bank workbook from either extractor script.")
    ap.add_argument("--into", default=None,
                    help="Existing Dönem x Banka workbook to append to "
                         "(its bank columns and quarters are kept).")
    ap.add_argument("--out", required=True)
    ap.add_argument("--add-bank", action="append", default=[], metavar="NAME",
                    help="Add a bank column not in --into. Repeatable.")
    ap.add_argument("--banks", default=None,
                    help="Comma-separated bank columns for a fresh build "
                         "(default: the report's 18).")
    args = ap.parse_args()

    src, src_quarters, src_banks = read_source(args.source)

    if args.into:
        data, quarters, banks = read_existing(args.into)
    else:
        data, quarters = {}, []
        banks = ([b.strip() for b in args.banks.split(",")]
                 if args.banks else list(DEFAULT_BANKS))

    for name in args.add_bank:
        if name not in banks:
            banks.append(name)
    order = {b: i for i, b in enumerate(CANONICAL_BANKS)}
    unknown = [b for b in banks if b not in order]
    if unknown:
        raise SystemExit("not in CANONICAL_BANKS: " + "; ".join(unknown))
    banks.sort(key=lambda b: order[b])

    missing = [b for b in banks if b not in src_banks]
    data.update(src)
    for q in src_quarters:
        if q not in quarters:
            quarters.append(q)
    quarters.sort(key=quarter_key)

    build(data, quarters, banks).save(args.out)

    print(f"Wrote {args.out}")
    print(f"  quarters : {len(quarters)}  ({quarters[0]} .. {quarters[-1]})")
    print(f"  banks    : {len(banks)}")
    print(f"  rows     : {len(quarters) * 10}")
    if src_quarters:
        print(f"  from source: {', '.join(src_quarters)}")
    if missing:
        print("  WARNING: no data in --source for, written as 0: "
              + "; ".join(missing))
    return 0


if __name__ == "__main__":
    sys.exit(main())
