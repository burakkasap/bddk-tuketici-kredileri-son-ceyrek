# BDDK Tüketici Kredileri — Konsolide Rapor Çıkarımı

Tools to extract the **consumer-loan breakdown** ("Tüketici kredileri, bireysel
kredi kartları ve personel kredi kartlarına ilişkin bilgiler") from Turkish banks'
**consolidated** independent audit reports published on BDDK's
[Bağımsız Denetim Raporları portal](https://www.bddk.gov.tr/BdrUyg), into an Excel
workbook laid out like `Template.xlsx`.

For each bank and quarter, ten fields are captured (thousand TRY):

```
Tüketici Kredileri-TP   (total)      Tüketici Kredileri-YP   (total)
  Konut Kredisi                        Konut Kredisi
  Taşıt Kredisi                        Taşıt Kredisi
  İhtiyaç Kredisi                      İhtiyaç Kredisi
  Diğer                                Diğer
```

Rules: no consolidated report for a bank/quarter → `0`; a blank / `-` field → `0`.

## Two tools

### 1. Historical backfill — [`bddk_consolidated_loans.py`](bddk_consolidated_loans.py)
One-off scrape of a quarter **range** across banks. Produced the workbooks in this repo:

- `BDDK_Konsolide_Tuketici_Kredileri.xlsx` — all 72 banks, 2023 Q1 – 2026 Q1.
- `BDDK_Konsolide_Tuketici_Kredileri_25Banks.xlsx` — the 25 banks that actually file
  consolidated consumer-loan data.

```bash
pip install -r requirements.txt
python3 bddk_consolidated_loans.py                 # full range, all banks
python3 bddk_consolidated_loans.py --only-banks    # just the 25 tracked banks
```

### 2. Recurring quarterly tool — [`bddk_quarterly/`](bddk_quarterly/)
Run each quarter with **no code changes** to pull **only the latest available quarter**
for the 25 tracked banks. Auto-detects the newest published quarter; adds a pip-only
**OCR fallback** (RapidOCR, no Tesseract/Homebrew) for banks that publish the table as
an image — every OCR value is gated by a sum-check
(`Konut + Taşıt + İhtiyaç + Diğer == the -TP / -YP total`).

```bash
cd bddk_quarterly
pip install -r requirements.txt
python3 bddk_latest_quarter.py                 # auto-detect newest quarter
python3 bddk_latest_quarter.py --quarter 2026Q2   # or pin one
```

See [`bddk_quarterly/README.md`](bddk_quarterly/README.md) for details.

## How the extraction is robust

Both tools read the PDF text layer and handle the many shapes BDDK reports come in:
zipped/raw-PDF/`.docx` archives, zip-within-zip, multi-file archives, case and
whitespace variants in table labels, and Turkish vs. US number formatting. The
quarterly tool additionally rasterizes + OCRs pages whose table is an image / vector
curves / letter-spaced text, accepting the numbers only when the internal sum-check
passes.

## Repo contents

| Path | What |
|---|---|
| `bddk_consolidated_loans.py` | Historical range backfill script |
| `bddk_quarterly/bddk_latest_quarter.py` | Recurring latest-quarter tool |
| `Template.xlsx` | Target output layout |
| `BDDK_Konsolide_Tuketici_Kredileri*.xlsx` | Extracted deliverables |
| `bddk_scrape*_log.jsonl` | Per-bank/quarter run logs (provenance) |

Downloaded report caches, OCR page renders, and per-run outputs are git-ignored
(`bddk_cache/` can be gigabytes and is regenerated on each run).

## Data source & disclaimer

All data is extracted from public consolidated audit reports on
[bddk.gov.tr](https://www.bddk.gov.tr/BdrUyg). This project is not affiliated with BDDK.
Figures are reproduced as reported (thousand TRY); verify against the source PDFs before
relying on them.
