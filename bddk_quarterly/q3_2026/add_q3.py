#!/usr/bin/env python3
"""Append the 2026 Q3 block to the user's Dönem x Banka workbook, in place.

Per-bank source is decided by the flag row in Sheet1:
  1 -> BDDK, read from the extractor workbook passed as --bddk
  2 -> the user's own report, read from q3_measured.json (million TL -> x1000)
  0 -> no report anywhere: grow Q2 by the bank's own Q1->Q2 growth rate

Styling, Sheet1, freeze panes and the TOPLAM formulas are preserved.
"""
import argparse
import copy
import json
import sys

import openpyxl

QUARTER = "2026 Q3"
PREV, PREV2 = "2026 Q2", "2026 Q1"

BLOCKS = [
    ("Tüketici", "Tüketici Kredileri (TP+YP)",
     ["Konut Kredisi", "Taşıt Kredisi", "İhtiyaç Kredisi", "Diğer"]),
    ("Ticari", "Taksitli Ticari Krediler (TP+YP)",
     ["İşyeri Kredisi", "Taşıt Kredisi", "İhtiyaç Kredisi", "Diğer"]),
]
# Sub-item order the extractor workbook writes, per table sheet.
SRC_SHEET = {"Tüketici": "Tüketici Kredileri", "Ticari": "Taksitli Ticari Krediler"}


def read_rows(ws, quarter, banks):
    """-> {(tipi, kategori): {bank: value}} for one quarter."""
    out = {}
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value != quarter:
            continue
        key = (ws.cell(r, 2).value, ws.cell(r, 3).value)
        out[key] = {b: (ws.cell(r, 4 + i).value or 0) for i, b in enumerate(banks)}
    return out


def read_bddk(path):
    """Extractor workbook -> {(bank, tipi, kategori): TP+YP value}."""
    wb = openpyxl.load_workbook(path, data_only=True)
    data = {}
    for tipi, total_label, subs in BLOCKS:
        sheet = SRC_SHEET[tipi]
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        qcols = [c for c in range(7, ws.max_column + 1) if ws.cell(1, c).value]
        if not qcols:
            continue
        col = qcols[0]
        src_tp = total_label.replace(" (TP+YP)", "-TP")
        src_yp = total_label.replace(" (TP+YP)", "-YP")
        labels = [src_tp] + subs + [src_yp] + subs
        bank, pos = None, 0
        for r in range(1, ws.max_row + 1):
            lab = ws.cell(r, 6).value
            if lab is None:
                continue
            lab = str(lab).strip()
            if bank and pos < len(labels) and lab == labels[pos]:
                cat = total_label if pos in (0, 5) else lab
                key = (bank, tipi, cat)
                data[key] = data.get(key, 0) + (ws.cell(r, col).value or 0)
                pos += 1
            else:
                bank, pos = lab, 0
    return data


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--book", required=True)
    ap.add_argument("--measured", required=True)
    ap.add_argument("--bddk", default=None,
                    help="Extractor workbook for the flag-1 banks. Omitted: they stay 0.")
    ap.add_argument("--out", required=True)
    ap.add_argument("--bddk-scale", type=int, default=1000,
                    help="Multiplier for the extractor's figures. The 2026 Q3 "
                         "reports all declare Milyon TL, so 1000 reaches Bin TL.")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.book)
    ws = wb["Dönem x Banka"]
    flags_ws = wb["Sheet1"]

    hdr = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    banks = [h for h in hdr[3:] if h and h != "TOPLAM"]
    ncol = 4 + len(banks)                       # TOPLAM column index
    flags = {hdr[c - 1]: flags_ws.cell(2, c).value for c in range(4, 4 + len(banks))}

    measured = json.load(open(args.measured, encoding="utf-8"))
    bddk = read_bddk(args.bddk) if args.bddk else {}

    q2 = read_rows(ws, PREV, banks)
    q1 = read_rows(ws, PREV2, banks)
    if not q2 or not q1:
        sys.exit("missing 2026 Q1/Q2 rows in the workbook")

    # Growth factor per flag-0 bank and table: the TOTAL row's Q1->Q2 ratio,
    # applied to every line of that block so sub-items still sum to the total.
    # The total is then set to the sum of the rounded sub-items, so rounding
    # cannot break the sum-check.
    grown = {}
    for tipi, total_label, subs in BLOCKS:
        for b in banks:
            a = q1[(tipi, total_label)][b]
            c = q2[(tipi, total_label)][b]
            rt = (c / a) if a else 0.0
            vals = {s: round(q2[(tipi, s)][b] * rt) for s in subs}
            vals[total_label] = sum(vals.values())
            grown[(b, tipi)] = vals

    # Style templates: the first quarter block's total row and a body row.
    first = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value)
    tmpl_total, tmpl_body = first, first + 1

    r = ws.max_row + 1
    report = []
    for tipi, total_label, subs in BLOCKS:
        for cat in [total_label] + subs:
            is_total = cat == total_label
            src_row = tmpl_total if is_total else tmpl_body
            ws.cell(r, 1, QUARTER)
            ws.cell(r, 2, tipi)
            ws.cell(r, 3, cat)
            for i, b in enumerate(banks):
                f = flags.get(b)
                if f == 2:
                    m = measured.get(b)
                    if m is None:          # report promised but not supplied yet
                        v, how = 0, "MISSING-report"
                    else:
                        key = "total" if is_total else cat
                        v = round(m[tipi][key] * (1000 if m["unit"] == "million" else 1))
                        how = "report"
                elif f == 1:
                    v = int(bddk.get((b, tipi, cat), 0) or 0) * args.bddk_scale
                    how = "bddk"
                else:
                    v = grown[(b, tipi)][cat]
                    how = "extrapolated"
                ws.cell(r, 4 + i, v)
                if is_total:
                    report.append((b, tipi, how, v))
            ws.cell(r, ncol, f"=SUM(D{r}:{openpyxl.utils.get_column_letter(ncol-1)}{r})")
            for c in range(1, ncol + 1):
                s = ws.cell(src_row, c)
                d = ws.cell(r, c)
                d.font = copy.copy(s.font)
                d.fill = copy.copy(s.fill)
                d.number_format = s.number_format
                d.alignment = copy.copy(s.alignment)
            r += 1

    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(ncol)}{r-1}"
    wb.save(args.out)

    # Sum-check every written block.
    chk = openpyxl.load_workbook(args.out, data_only=True)["Dönem x Banka"]
    rows = {}
    for rr in range(2, chk.max_row + 1):
        if chk.cell(rr, 1).value == QUARTER:
            rows[(chk.cell(rr, 2).value, chk.cell(rr, 3).value)] = \
                [chk.cell(rr, 4 + i).value or 0 for i in range(len(banks))]
    bad = 0
    for tipi, total_label, subs in BLOCKS:
        for i, b in enumerate(banks):
            s = sum(rows[(tipi, x)][i] for x in subs)
            t = rows[(tipi, total_label)][i]
            if s != t:
                bad += 1
                print(f"  SUM-CHECK FAIL  {b} / {tipi}: subs={s} total={t}")
    print(f"Wrote {args.out}  ({QUARTER} added, {len(banks)} banks)")
    print(f"sum-check: {len(banks)*2 - bad}/{len(banks)*2} blocks pass")
    print("\nper-bank source (total rows):")
    for b, tipi, how, v in report:
        print(f"  {b:<33} {tipi:<9} {how:<13} {v:>15,}")


if __name__ == "__main__":
    main()
