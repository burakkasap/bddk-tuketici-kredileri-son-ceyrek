# BDDK Quarterly Loan-Breakdown Extractor (latest quarter, 25 banks)

Each quarter, re-run this tool to pull **only the latest available quarter**'s
consolidated loan breakdowns for the **25 tracked banks**, in the Template.xlsx
layout. **You never edit the code between quarters** — when 2026-Q2 is published,
just run it again and it picks 2026-Q2 up automatically; same for 2026-Q3.

It extracts **two tables**, written as **two sheets** of one workbook (their
sub-items differ, so they are kept separate):

| sheet `Tüketici Kredileri` | sheet `Taksitli Ticari Krediler` |
|---|---|
| Tüketici Kredileri-**TP** | Taksitli Ticari Krediler-**TP** |
| &nbsp;&nbsp;Konut / Taşıt / İhtiyaç / Diğer | &nbsp;&nbsp;**İşyeri** / Taşıt / İhtiyaç / Diğer |
| Tüketici Kredileri-**YP** | Taksitli Ticari Krediler-**YP** |
| &nbsp;&nbsp;(same four) | &nbsp;&nbsp;(same four) |

Only the **Cari Dönem** (current period) figures are used — every report also
contains an Önceki Dönem (previous period) copy, which is explicitly excluded.

## One-time setup

```bash
cd bddk_quarterly
python3 -m pip install -r requirements.txt
```

No Tesseract or Homebrew needed — OCR (used only when a bank publishes its table
as an image) runs via the pip-only `rapidocr-onnxruntime`.

## Run it

```bash
# Auto-detect and extract the newest published quarter (both tables):
python3 bddk_latest_quarter.py

# Pin a specific quarter (no code edit needed):
python3 bddk_latest_quarter.py --quarter 2026Q2

# Only one table, if you ever want a single sheet:
python3 bddk_latest_quarter.py --tables taksitli
```

Output (quarter encoded in the filenames, so runs never overwrite each other):

```
output/BDDK_25Banks_2026Q2.xlsx   # two sheets, one column each: the quarter's values
output/run_2026Q2.jsonl           # per-bank/table status log
```

## What each run does

1. Scans the current and previous year for all 25 banks and picks the **newest
   quarter any of them has filed**. It prints how many of the 25 have a report and
   **warns if fewer than half** have filed yet (filing is staggered — if you ran too
   early, re-run later or pass `--quarter`).
2. Downloads each bank's consolidated report once (cached under `bddk_cache/`) and
   extracts **both** tables from it.
3. Per bank/table it tries, in order:
   - **verified overrides** (a few known-pathological reports),
   - **text / .docx parsing** (the normal path),
   - **OCR fallback** — if the table is an image / vector curves / letter-spaced
     text, or if the text parse's own arithmetic doesn't hold (some PDFs inject
     spaces inside numbers), the page is rasterized and read with RapidOCR.
4. **Every figure is validated by the report's own arithmetic**
   (`sub-items == the -TP / -YP total`). Anything that can't be validated is left 0
   and flagged rather than written as an unverified number.

## Reading the summary

- `text-parsed` / `OCR` / `override` — populated and sum-check-validated.
- `no-note(0)` / `banks with no report` — legitimately 0 (bank didn't file, or the
  report has no such note).
- **NEEDS MANUAL REVIEW** — could not be validated; cell left 0 and the page image
  saved to `_ocr_pages/`. Open that PNG, read the ~10 numbers, and add an entry to
  `MANUAL_OVERRIDES` at the top of the script (it is sum-checked at startup).
- **ERRORS** — usually a file missing/broken on BDDK's own server; left as 0.

## Rules

- No consolidated report for the quarter → all fields `0`.
- Blank / `-` field in the report → `0`.

## Notes

- Self-contained and independent of the historical backfill
  (`../bddk_consolidated_loans.py`); the two share the same table registry and
  parsing rules.
- `bddk_cache/`, `_ocr_pages/` and `output/` are created on first run and are
  git-ignored.
