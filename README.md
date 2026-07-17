# BDDK Kredi Dipnotları — Konsolide Rapor Çıkarımı

Tools to extract loan-breakdown footnote tables from Turkish banks' **consolidated**
independent audit reports published on BDDK's
[Bağımsız Denetim Raporları portal](https://www.bddk.gov.tr/BdrUyg), into Excel
workbooks laid out like `Template.xlsx`.

**Two tables** are extracted, kept as **two separate sheets** (their sub-items differ):

| sheet `Tüketici Kredileri` | sheet `Taksitli Ticari Krediler` |
|---|---|
| Tüketici Kredileri-**TP** / -**YP** | Taksitli Ticari Krediler-**TP** / -**YP** |
| &nbsp;&nbsp;**Konut** Kredisi | &nbsp;&nbsp;**İşyeri** Kredisi |
| &nbsp;&nbsp;Taşıt Kredisi | &nbsp;&nbsp;Taşıt Kredisi |
| &nbsp;&nbsp;İhtiyaç Kredisi | &nbsp;&nbsp;İhtiyaç Kredisi |
| &nbsp;&nbsp;Diğer | &nbsp;&nbsp;Diğer |

Figures are the **Toplam** column, in thousand TRY. Only the **Cari Dönem** (current
period) table is ever used — every report also contains an Önceki Dönem (previous
period) copy, which is explicitly excluded.

Rules: no consolidated report for a bank/quarter → `0`; blank / `-` field → `0`.

## Two tools

### 1. Historical backfill — [`bddk_consolidated_loans.py`](bddk_consolidated_loans.py)

One-off scrape of a quarter **range**. Produced the workbooks in this repo:

- `BDDK_Konsolide_25Banks_2023Q1-2026Q1.xlsx` — **both tables**, 25 tracked banks,
  2023 Q1 – 2026 Q1 (two sheets).
- `BDDK_Konsolide_Tuketici_Kredileri.xlsx` — earlier run: all 72 banks, Tüketici only.
- `BDDK_Konsolide_Tuketici_Kredileri_25Banks.xlsx` — earlier run: 25 banks, Tüketici only.
  Superseded by the two-sheet workbook above, which also corrects several figures
  (see *Data-quality fixes*).

```bash
pip install -r requirements.txt
python3 bddk_consolidated_loans.py --only-banks              # 25 banks, both tables
python3 bddk_consolidated_loans.py --tables taksitli         # one sheet only
python3 bddk_consolidated_loans.py --start 2023Q1 --end 2026Q1
```

### 2. Recurring quarterly tool — [`bddk_quarterly/`](bddk_quarterly/)

Run each quarter with **no code changes** to pull **only the latest available quarter**
for the 25 tracked banks. Auto-detects the newest published quarter.

```bash
cd bddk_quarterly
pip install -r requirements.txt
python3 bddk_latest_quarter.py                    # auto-detect newest quarter
python3 bddk_latest_quarter.py --quarter 2026Q2   # or pin one
```

See [`bddk_quarterly/README.md`](bddk_quarterly/README.md) for details.

## How the extraction stays honest

Every figure is validated by **the report's own arithmetic**:
`Konut/İşyeri + Taşıt + İhtiyaç + Diğer == the -TP / -YP total`. Anything that can't be
validated is left `0` and flagged with a page image for review — never written as an
unverified number. In the shipped workbook, **all 625 populated cells pass this check.**

Both tools read the PDF text layer and handle the many shapes BDDK reports come in:
zipped / raw-PDF / `.docx` archives, zip-within-zip, multi-file archives, Turkish vs US
number formats, label variants (`Krediler` vs `Kredileri`, `İş Yeri` vs `İşyeri`,
HSBC's `Otomobil Kredisi`), tight kerning that swallows spaces, footnote digits inside
labels, and figures wrapped onto the next line. When a table isn't in the text layer at
all — an image, vector curves, or letter-spaced text — the page is rasterized and read
with **RapidOCR** (pip-only; no Tesseract/Homebrew), again gated by the sum-check.

## Data-quality fixes found by the sum-check

Building the Taksitli table surfaced real errors in the earlier Tüketici-only workbook,
all corrected in the two-sheet workbook:

| Bank | Was | Now |
|---|---|---|
| HSBC (13 quarters) | `Taşıt Kredisi` = 0 — the bank labels it **`Otomobil Kredisi`** | actual figures |
| Türkiye Finans (9 quarters) | **previous-period**, digit-mangled figures (its Cari table is letter-spaced, so the text layer only exposed the Önceki one) | Cari figures via OCR |
| Emlak Katılım (2 quarters) | footnote digit inside a label (`Kredis4i`); one PDF draws every figure a row below its label | corrected / verified override |
| Vakıf Katılım (1 quarter) | `-TP` total blank — its figures wrap onto the next line | actual figures |

## Repo contents

| Path | What |
|---|---|
| `bddk_consolidated_loans.py` | Historical range backfill (both tables) |
| `bddk_quarterly/bddk_latest_quarter.py` | Recurring latest-quarter tool |
| `Template.xlsx` | Target output layout |
| `BDDK_Konsolide_25Banks_2023Q1-2026Q1.xlsx` | **Main deliverable** — both tables, two sheets |
| `BDDK_Konsolide_Tuketici_Kredileri*.xlsx` | Earlier Tüketici-only runs |
| `bddk_*_log.jsonl` | Per-bank/quarter/table run logs (provenance) |

Downloaded report caches, OCR page renders and per-run outputs are git-ignored
(`bddk_cache/` can be gigabytes and is regenerated on each run).

## Data source & disclaimer

All data is extracted from public consolidated audit reports on
[bddk.gov.tr](https://www.bddk.gov.tr/BdrUyg). This project is not affiliated with BDDK.
Figures are reproduced as reported (thousand TRY); verify against the source PDFs before
relying on them.
