import re
import time
import pytesseract
from pdf2image import convert_from_path
from PIL import ImageEnhance
import pypdf
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# SETTINGS — edit these before running
# ─────────────────────────────────────────────────────────────────────────────

PDF_PATH    = r"C:\Users\YourName\Documents\tender.pdf"   # ← your input PDF
OUTPUT_PATH = r"C:\Users\YourName\Documents\tender_extracted.xlsx"  # ← output Excel

DPI = 300   # 300 = best quality; 200 = faster (~2× speed, slightly less accurate)

# ─────────────────────────────────────────────────────────────────────────────
# COLUMN CALIBRATION
# These fractional positions mark the RIGHT edge of each column,
# measured as a fraction of page width.
# Calibrated for the standard A4 Schedule of Rates table layout.
# Run the calibration snippet at the bottom of this file if your PDF differs.
# ─────────────────────────────────────────────────────────────────────────────

COL_SEPS_FRAC = [0.1407, 0.5377, 0.6183, 0.6989, 0.7997]
COL_NAMES     = ["item", "desc", "unit", "qty", "rate", "amount"]

TABLE_Y0    = 0.10   # fraction of page height where table body starts
TABLE_Y1    = 0.87   # fraction of page height where table body ends
OCR_CONF    = 20     # minimum Tesseract confidence (0–100)
ROW_SNAP_PX = 130    # max Y distance (px at 300 DPI) to snap a value to an item row

COL_NOISE = {
    "item":   {"item", "i", "te", "ter", "tem"},
    "unit":   {"unit"},
    "qty":    {"qty", "quantity"},
    "rate":   {"rate"},
    "amount": {"amount", "amount _|", "|"},
}

VALID_UNITS = {"item", "nr", "no", "m2", "m3", "m", "ls", "lsum",
               "set", "lot", "run", "%", "pc", "pcs", "roll", "-"}

ITEM_RE      = re.compile(r"^[A-Z]$|^\d{1,3}$")
SCHEDULE_RE  = re.compile(r"SCHEDULE\s+NO[.\s]+([0-9.]+)", re.IGNORECASE)
PAGE_REF_RE  = re.compile(r"\bS\s*\d+[\./]\d+\b")
DESC_NOISE   = re.compile(
    r"^(SCHEDULE\s+NO\.|JSSC\s+|WORKS\s+AND|BUILDING\s+\d|^Description$|"
    r"TOTAL\s+CARRIED|\(Qn\.|^Notes?[:\-]|^SERVICES$)",
    re.IGNORECASE,
)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def fix_item_label(raw):
    t = raw.upper().strip()
    if len(t) == 2 and t[0] == t[1]:          return t[0]   # Cc → C
    if len(t) == 2 and t[1] in ".,)'|":       return t[0]   # A. → A
    misreads = {"5)": "C", "8)": "B", "0)": "O"}
    return misreads.get(t, t if len(t) == 1 else "")

def is_num_or_dash(s):
    return bool(re.match(r"^[\d,.\-–]+$", s.strip()))

def is_cid_garbled(text):
    return text.count("(cid:") > 3

# ─────────────────────────────────────────────────────────────────────────────
# PAGE METADATA  (schedule number + page reference)
# ─────────────────────────────────────────────────────────────────────────────

def page_metadata(img):
    small = img.resize((img.width // 3, img.height // 3))
    text  = pytesseract.image_to_string(small, config="--psm 6")
    sch   = (m.group(1) if (m := SCHEDULE_RE.search(text)) else "")
    refs  = PAGE_REF_RE.findall(text)
    ref   = refs[-1].replace(" ", "") if refs else ""
    return sch, ref

# ─────────────────────────────────────────────────────────────────────────────
# PER-PAGE EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

def process_page(img):
    img_e = ImageEnhance.Contrast(img).enhance(1.8)
    w, h  = img_e.size

    schedule_no, page_ref = page_metadata(img)

    seps_px = [int(f * w) for f in COL_SEPS_FRAC] + [w]

    def col_of(x):
        for name, sep in zip(COL_NAMES, seps_px):
            if x < sep:
                return name
        return "amount"

    # get all word bounding boxes via Tesseract
    df = pytesseract.image_to_data(img_e,
                                   output_type=pytesseract.Output.DATAFRAME,
                                   config="--psm 6")
    df = df[(df["conf"] > OCR_CONF) & (df["text"].str.strip().str.len() > 0)].copy()

    Y0_PX, Y1_PX = int(TABLE_Y0 * h), int(TABLE_Y1 * h)
    body = df[(df["top"] > Y0_PX) & (df["top"] < Y1_PX)].copy()
    body["col"] = body["left"].apply(col_of)

    # --- find item labels (A, B, C … or 1, 2, 3 …) -----------------------
    items_raw  = body[body["col"] == "item"]
    item_rows  = []
    seen       = set()

    for _, ir in items_raw.sort_values("top").iterrows():
        label = fix_item_label(str(ir["text"]))
        if not ITEM_RE.match(label) or label in seen:
            continue
        seen.add(label)
        item_rows.append({"item": label, "item_y": int(ir["top"])})

    if not item_rows:
        return schedule_no, page_ref, []

    item_ys = [r["item_y"] for r in item_rows]

    # --- snap unit / qty / rate / amount to nearest item row --------------
    for col_name in ("unit", "qty", "rate", "amount"):
        col_words = body[body["col"] == col_name].copy()
        noise     = COL_NOISE.get(col_name, set())
        col_words = col_words[~col_words["text"].str.lower().str.strip().isin(noise)]

        for row in item_rows:
            near   = col_words[((col_words["top"] - row["item_y"]).abs() < ROW_SNAP_PX)]
            tokens = near.sort_values("left")["text"].tolist()
            row[col_name] = " ".join(tokens)

    # --- build description for each item row ------------------------------
    desc_words = body[body["col"] == "desc"].sort_values("top")

    for i, row in enumerate(item_rows):
        iy      = row["item_y"]
        y_start = iy - 70
        y_end   = item_ys[i + 1] - 70 if i + 1 < len(item_ys) else Y1_PX
        seg     = desc_words[(desc_words["top"] >= y_start) & (desc_words["top"] < y_end)]

        # reconstruct lines preserving reading order
        seg_s   = seg.sort_values(["top", "left"])
        lines   = []
        prev_top, buf = None, []
        for _, wd in seg_s.iterrows():
            if prev_top is None or abs(wd["top"] - prev_top) > 8:
                if buf: lines.append(" ".join(buf))
                buf, prev_top = [wd["text"]], wd["top"]
            else:
                buf.append(wd["text"])
        if buf: lines.append(" ".join(buf))

        lines = [l for l in lines if not DESC_NOISE.match(l)]
        desc  = " ".join(lines).strip()
        desc  = re.sub(r"\s*\(?Cont'?d\)?", "", desc, flags=re.IGNORECASE).strip()
        row["desc"] = desc

    # --- detect section heading (first ALL-CAPS line above first item) ----
    pre_words = desc_words[desc_words["top"] < (item_ys[0] - 70)]
    section   = ""
    for _, wd in pre_words.sort_values("top").iterrows():
        t = wd["text"].strip()
        if t.isupper() and len(t) > 4:
            section = t
            break

    # --- assemble final rows ----------------------------------------------
    rows = []
    for row in item_rows:
        unit = row.get("unit", "").lower().strip()
        qty  = row.get("qty",  "").strip()
        rate = row.get("rate", "").strip()
        amt  = row.get("amount", "").strip()

        rows.append({
            "Item":        row["item"],
            "Section":     section,
            "Description": row.get("desc", ""),
            "Unit":        unit if unit in VALID_UNITS else "",
            "Qty":         qty  if is_num_or_dash(qty)  else "",
            "Rate":        rate if is_num_or_dash(rate) else "",
            "Amount":      amt  if is_num_or_dash(amt)  else "",
        })

    return schedule_no, page_ref, rows

# ─────────────────────────────────────────────────────────────────────────────
# FULL PDF PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────

def process_pdf(pdf_path, output_path):
    reader      = pypdf.PdfReader(pdf_path)
    total_pages = len(reader.pages)
    all_data    = []

    # try to get project name from text layer
    project = ""
    try:
        text = reader.pages[0].extract_text() or ""
        if not is_cid_garbled(text):
            for line in text.split("\n"):
                line = line.strip()
                pc   = sum(32 <= ord(c) < 127 for c in line)
                if len(line) > 15 and pc / max(len(line), 1) > 0.85:
                    project = line
                    break
    except Exception:
        pass

    print(f"Starting extraction: {total_pages} pages  |  DPI={DPI}", flush=True)
    if project:
        print(f"Project: {project}", flush=True)
    print("-" * 55, flush=True)
    start = time.time()

    for i in range(total_pages):
        elapsed = time.time() - start
        print(f"Page {i+1}/{total_pages}  [{elapsed:.1f}s]  ", end="", flush=True)

        try:
            images = convert_from_path(pdf_path, dpi=DPI,
                                       first_page=i+1, last_page=i+1)
            if not images:
                print("(no image)")
                continue

            sch, ref, rows = process_page(images[0])

            for row in rows:
                row["Project"]     = project
                row["Schedule No"] = sch
                row["Page"]        = i + 1
                row["Page Ref"]    = ref
                all_data.append(row)

            print(f"{len(rows)} items  (Sch {sch or '?'}, {ref or '?'})", flush=True)

        except Exception as e:
            print(f"ERROR — {e}", flush=True)

    # save intermediate CSV in case Excel write fails
    df = pd.DataFrame(all_data, columns=[
        "Project", "Schedule No", "Page", "Page Ref",
        "Item", "Section", "Description", "Unit", "Qty", "Rate", "Amount"
    ])
    csv_path = output_path.replace(".xlsx", ".csv")
    df.to_csv(csv_path, index=False)
    print(f"\nCSV saved: {csv_path}", flush=True)

    # write formatted Excel
    save_excel(df, output_path)

    elapsed = time.time() - start
    print(f"\nDone. {len(df)} rows extracted in {elapsed:.1f}s", flush=True)


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

DARK_BLUE = "1F3864"
MID_BLUE  = "2E75B6"
ALT_ROW   = "EBF3FB"
_S = Side(style="thin", color="BFBFBF")
TB = Border(left=_S, right=_S, top=_S, bottom=_S)

def _c(ws, r, c, v, bold=False, fg="000000", bg=None, ha="left", wrap=False, sz=9):
    cell            = ws.cell(r, c, v)
    cell.font       = Font(bold=bold, color=fg, size=sz)
    cell.alignment  = Alignment(horizontal=ha, vertical="top", wrap_text=wrap)
    cell.border     = TB
    if bg:
        cell.fill   = PatternFill("solid", fgColor=bg)
    return cell

def save_excel(df, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Schedule of Rates"

    COLS = [
        ("Project",     36, "left"),
        ("Schedule No", 13, "center"),
        ("Page",         6, "center"),
        ("Page Ref",    10, "center"),
        ("Item",         6, "center"),
        ("Section",     26, "left"),
        ("Description", 56, "left"),
        ("Unit",         8, "center"),
        ("Qty",          8, "center"),
        ("Rate",        13, "right"),
        ("Amount",      13, "right"),
    ]

    for ci, (name, w, _) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # header row
    ws.row_dimensions[1].height = 30
    for ci, (name, _, _) in enumerate(COLS, 1):
        _c(ws, 1, ci, name, bold=True, fg="FFFFFF", bg=DARK_BLUE, ha="center", sz=10)

    # data rows
    dr, prev_sch = 2, None
    for _, rec in df.iterrows():
        sch = rec.get("Schedule No", "")

        # blue banner when schedule number changes
        if sch and sch != prev_sch:
            ws.merge_cells(start_row=dr, start_column=1,
                           end_row=dr,   end_column=len(COLS))
            banner            = ws.cell(dr, 1, f"  SCHEDULE NO. {sch}")
            banner.font       = Font(bold=True, color="FFFFFF", size=10)
            banner.fill       = PatternFill("solid", fgColor=MID_BLUE)
            banner.alignment  = Alignment(horizontal="left", vertical="center")
            banner.border     = TB
            ws.row_dimensions[dr].height = 22
            dr      += 1
            prev_sch = sch

        bg   = ALT_ROW if dr % 2 == 0 else "FFFFFF"
        vals = [rec.get(col[0], "") for col in COLS]
        for ci, (v, (_, _, ha)) in enumerate(zip(vals, COLS), 1):
            _c(ws, dr, ci, v, bg=bg, ha=ha, wrap=(ci == 7))
        ws.row_dimensions[dr].height = 28
        dr += 1

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}1"

    # summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 24
    rows_s = [
        ("Field",            "Value"),
        ("Total Line Items",  len(df)),
        ("Schedules Found",   ", ".join(df["Schedule No"].dropna().unique()) or "N/A"),
        ("Pages with Items",  int(df["Page"].nunique())),
        ("Items with Qty",    int((df["Qty"].str.strip().replace({"-": "", "–": ""}) != "").sum())),
        ("Input PDF",         PDF_PATH),
        ("Output",            output_path),
    ]
    for r, (label, val) in enumerate(rows_s, 1):
        ws2.cell(r, 1, label).font = Font(bold=(r==1),
                                          color="FFFFFF" if r==1 else "000000", size=10)
        ws2.cell(r, 2, str(val)).font = Font(size=10)
        if r == 1:
            for c in (ws2.cell(r,1), ws2.cell(r,2)):
                c.fill      = PatternFill("solid", fgColor=DARK_BLUE)
                c.alignment = Alignment(horizontal="center")

    wb.save(output_path)
    print(f"Excel saved: {output_path}", flush=True)


# ─────────────────────────────────────────────────────────────────────────────
# RUN
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    process_pdf(PDF_PATH, OUTPUT_PATH)


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN CALIBRATION SNIPPET
# If your PDF uses different column widths, run this first to find COL_SEPS_FRAC.
# 1. Paste your PDF path into pdf_path below
# 2. Run this block once
# 3. Copy the printed COL_SEPS_FRAC value into the SETTINGS section above
# ─────────────────────────────────────────────────────────────────────────────
#
# import numpy as np
# from pdf2image import convert_from_path
#
# pdf_path = r"C:\path\to\your.pdf"
# imgs = convert_from_path(pdf_path, dpi=300, first_page=1, last_page=1)
# img  = imgs[0]
# w, h = img.size
# arr  = np.array(img.convert("L"))
#
# header_band = arr[int(0.09*h):int(0.12*h), :]
# col_min     = header_band.min(axis=0)
# dark_cols   = [i for i in range(5, w-5) if col_min[i] < 50]
#
# clusters, cl = [], [dark_cols[0]]
# for p in dark_cols[1:]:
#     if p - cl[-1] <= 5: cl.append(p)
#     else:
#         clusters.append(int(sum(cl)/len(cl))); cl = [p]
# clusters.append(int(sum(cl)/len(cl)))
#
# print("COL_SEPS_FRAC =", [round(c/w, 4) for c in clusters])
