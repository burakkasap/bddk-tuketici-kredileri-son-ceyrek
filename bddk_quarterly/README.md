# BDDK Quarterly Consumer-Loan Extractor (latest quarter, 25 banks)

Each quarter, re-run this tool to pull **only the latest available quarter**'s
consolidated consumer-loan breakdown for the **25 tracked banks**, in the
Template.xlsx layout. **You never edit the code between quarters** — when
2026-Q2 is published, just run it again and it picks 2026-Q2 up automatically;
same for 2026-Q3, and so on.

## One-time setup

```bash
cd bddk_quarterly
python3 -m pip install -r requirements.txt
```

No Tesseract or Homebrew needed — OCR (used only when a bank publishes its
table as an image instead of text) runs via the pip-only `rapidocr-onnxruntime`.

## Run it

```bash
# Auto-detect and extract the newest published quarter:
python3 bddk_latest_quarter.py

# Pin a specific quarter (no code edit needed):
python3 bddk_latest_quarter.py --quarter 2026Q2
```

Output (quarter encoded in the filenames, so runs never overwrite each other):

```
output/BDDK_25Banks_2026Q2.xlsx   # one column: the quarter's values, 25 banks
output/run_2026Q2.jsonl           # per-bank status log
```

## What each run does

1. Scans the current and previous year for all 25 banks and picks the **newest
   quarter any of them has filed**. It prints how many of the 25 have a report
   and **warns if fewer than half** have filed yet (filing is staggered — if you
   ran too early, re-run later or pass `--quarter`).
2. Downloads each bank's consolidated report for that quarter (cached under
   `bddk_cache/`, so re-runs don't re-download).
3. Extracts the 10 fields per bank (Tüketici Kredileri-TP total + Konut / Taşıt /
   İhtiyaç / Diğer, and the same for -YP), trying, in order:
   - **verified overrides** (a few known historical cells),
   - **text / .docx parsing** (the normal path),
   - **OCR fallback** — if a bank renders the table as an image / vector curves /
     letter-spaced text (e.g. Fibabanka), the page is rasterized and read with
     RapidOCR. OCR numbers are accepted **only if they pass the sum-check**
     (Konut + Taşıt + İhtiyaç + Diğer = the -TP total, same for -YP).
4. Writes the workbook and a summary.

## Reading the summary

- `text-parsed` / `OCR` / `override` — successfully populated (OCR values are
  sum-check-validated).
- `no-report(0)` — that bank has no consolidated report for the quarter → 0
  (expected; e.g. a bank that hadn't started consolidated filing yet).
- **NEEDS MANUAL REVIEW** — the table wasn't text-readable *and* OCR couldn't be
  trusted (sum-check failed, or the OCR engine isn't installed). The cell is left
  0 and the page image is saved to `_ocr_pages/`. Open that PNG, read the ~10
  numbers, and either correct the cell by hand or add a `(eft, year, month)`
  entry to `MANUAL_OVERRIDES` at the top of the script.
- **ERRORS** — usually a file that is missing/broken on BDDK's own server
  (returns an HTML error page instead of the report); left as 0.

## Rules

- No consolidated report for the quarter → all 10 fields `0`.
- Blank / `-` field in the report → `0`.

## Notes

- This folder is self-contained and independent of the historical backfill
  (`../bddk_consolidated_loans.py`). Nothing here modifies the historical files.
- `bddk_cache/` and `_ocr_pages/` are created on first run.
