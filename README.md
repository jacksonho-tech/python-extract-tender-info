# Tender Extractor — Schedule of Rates PDF → Excel

A Python tool that extracts **Schedule of Rates** line items from construction tender PDFs into a structured Excel workbook.

---

## What It Does

Construction tender documents typically contain a Schedule of Rates table with columns:

| Item | Description | Unit | Qty | Rate | Amount |
|------|-------------|------|-----|------|--------|
| A    | foil shielded twisted pair cable ; two pair | item | - | - | |
| B    | low smoke zero halogen (LSZH) type ; 1.5mm2 | item | - | - | |

This tool reads the PDF, extracts each line item, and writes a clean Excel file with:
- **"Schedule of Rates"** sheet — one row per line item
- **"Summary"** sheet — extraction statistics

### Why OCR?

Many tender PDFs (especially those produced from older CAD/BIM workflows) use **broken custom font encodings**. Standard text extraction (pypdf, pdfplumber) returns garbled characters like `(cid:1)(cid:2)...`. This tool detects that automatically and falls back to **column-split OCR** using Tesseract — rasterising each page at 300 DPI then OCR-ing each column band separately.

---

## Installation

### Requirements

- Python 3.10+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) installed on your system
- [Poppler](https://poppler.freedesktop.org/) (for `pdf2image`)

**Ubuntu / Debian:**
```bash
sudo apt-get install tesseract-ocr poppler-utils
```

**macOS (Homebrew):**
```bash
brew install tesseract poppler
```

**Windows:**
- Tesseract: https://github.com/UB-Mannheim/tesseract/wiki
- Poppler: https://github.com/oschwartz10612/poppler-windows/releases

### Python packages

```bash
pip install -r requirements.txt
```

---

## Usage

```bash
# Basic usage — output named automatically
python extract_tender.py path/to/tender.pdf

# Specify output path
python extract_tender.py tender.pdf output.xlsx

# Faster (lower quality) — use 200 DPI
python extract_tender.py tender.pdf --dpi 200

# Suppress progress output
python extract_tender.py tender.pdf --quiet

# Print CSV to stdout instead of writing Excel
python extract_tender.py tender.pdf --no-excel
```

### Output columns

| Column | Description |
|--------|-------------|
| Project | Project name (from PDF metadata) |
| Schedule No | Schedule number (e.g. `4.7`) |
| Page | PDF page number |
| Page Ref | Page reference code (e.g. `S4.7/2`) |
| Item | Line item label (A, B, C … or 1, 2, 3 …) |
| Section | Section heading above the item |
| Description | Full item description |
| Unit | Unit of measurement (item, nr, m2, etc.) |
| Qty | Quantity |
| Rate | Rate (blank = to be filled by tenderer) |
| Amount | Amount (blank = to be filled by tenderer) |

---

## Column Calibration

The tool uses fractional column positions calibrated for the **standard A4 SOR table layout** (Item / Description / Unit / Qty / Rate / Amount). If your PDF uses different column widths, re-calibrate:

```python
# Run once on your PDF to find column separator positions
import numpy as np
from pdf2image import convert_from_path

imgs = convert_from_path("your.pdf", dpi=300, first_page=1, last_page=1)
img  = imgs[0]
w, h = img.size
arr  = np.array(img.convert("L"))

# Scan the header band for dark vertical lines
header_band = arr[int(0.09*h):int(0.12*h), :]
col_min = header_band.min(axis=0)
dark_cols = [i for i in range(5, w-5) if col_min[i] < 50]

# Cluster contiguous pixels into separator centres
clusters, cl = [], [dark_cols[0]]
for p in dark_cols[1:]:
    if p - cl[-1] <= 5: cl.append(p)
    else:
        clusters.append(int(sum(cl)/len(cl))); cl = [p]
clusters.append(int(sum(cl)/len(cl)))

print("COL_SEPS_FRAC =", [round(c/w, 4) for c in clusters])
```

Then update `COL_SEPS_FRAC` near the top of `extract_tender.py`.

---

## Project Structure

```
tender-extractor/
├── extract_tender.py       # Main extraction script
├── requirements.txt        # Python dependencies
├── README.md               # This file
├── .gitignore
└── sample/
    ├── tender.pdf          # Sample input
    └── tender_extracted.xlsx  # Sample output
```

---

## How It Works

1. **Font detection** — checks if pypdf returns CID-encoded garbage
2. **Rasterisation** — converts each page to a 300 DPI image via `pdf2image`
3. **Word bounding boxes** — runs `pytesseract.image_to_data()` to get every word's pixel coordinates
4. **Column assignment** — maps each word to a column (Item / Description / Unit / Qty / Rate / Amount) using the known vertical separator positions
5. **Row alignment** — identifies item labels (A, B, C…) in the Item column, then snaps unit/qty/rate/amount values to the nearest item by Y coordinate
6. **Description reconstruction** — collects Description-column words between each item's Y position and the next, preserving reading order
7. **Excel output** — writes a formatted workbook with schedule banner rows, alternating row colours, freeze pane, and auto-filter

---

## Limitations

- **Rate / Amount columns** — these are typically left blank in tender documents (filled in by the tenderer). The tool captures them when present.
- **Multi-line item descriptions** — reconstructed from word bounding boxes; word order is generally correct but may occasionally differ from the original for very complex layouts.
- **Non-standard layouts** — if your PDF uses significantly different column proportions, use the calibration snippet above to update `COL_SEPS_FRAC`.
- **Processing speed** — OCR at 300 DPI takes ~12–15 seconds per page. Use `--dpi 200` for faster (slightly less accurate) processing.


---

## License

MIT
