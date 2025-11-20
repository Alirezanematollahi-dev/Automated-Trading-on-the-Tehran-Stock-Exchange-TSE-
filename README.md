# Automated-Trading-on-the-Tehran-Stock-Exchange-TSE-
### An algorithmic trading system for Iran’s stock market that analyzes financial data, reports, and trading volumes to generate signals and support strategic integration.
## This script organizes and cleans Excel files containing financial statements downloaded from the Codal system. 
-- coding: utf-8 --¶
""" Split Excel into sections so that EACH heading + its content go to ONE sheet.

Rules:

Prefer banner: (dash row) + (TITLE row) + (dash row). Section starts at TITLE row (middle).
Otherwise detect "true title" rows (near-empty row whose text matches known titles or 'خلاصه ... - صفحه N') and the table starts soon (>=2 filled cells within a few rows).
Titles are normalized (ي/ی, ك/ک, spaces) so variant spellings are treated as the same logical title.
No duplicate sections for the same heading (dedup by normalized title + start proximity).
Copy preserves merges, styles, number formats, and column widths.
Output sheets are Right-to-Left.
Deps: pandas, openpyxl, tkinter """

import os, re from copy import copy import pandas as pd from openpyxl import load_workbook, Workbook from openpyxl.utils import get_column_letter from tkinter import Tk from tkinter.filedialog import askopenfilename

---------------- Config ----------------
KNOWN_TITLES = [ "صورت سود و زیان جامع تلفیقی", "صورت سود و زیان تلفیقی", "صورت وضعیت مالی تلفیقی", "صورت تغییرات در حقوق مالکانه تلفیقی", "صورت جریان های نقدی تلفیقی", "صورت سود و زیان", "صورت سود و زیان جامع", "صورت وضعیت مالی", "صورت تغییرات در حقوق مالکانه", "صورت جریان های نقدی", "اعضاي هيئت مديره", "شرکت های وابسته", "شرکت های فرعی", ] SUMMARY_REGEX = r"^خلاصه\s+اطلاعات\s+گزارش\s+تفسیری\s*-\s*صفحه\s+\d+$" TABLE_LOOKAHEAD = 12 # تا این تعداد ردیف پایین‌تر دنبال شروع جدول می‌گردیم DEDUP_PROXIMITY = 3 # اگر دو شروع با همین عنوان در <=3 ردیف همدیگه باشن، یکی حساب می‌شن

---------------- Helpers ----------------
def _norm(s): if not isinstance(s, str): return "" s = s.replace("ي","ی").replace("ك","ک").replace("\u0640","").replace("\u200c"," ") s = re.sub(r"\s+"," ", s).strip() return s

def _row_text(vals): parts = [str(v).strip() for v in vals if isinstance(v, str) and v and str(v).strip()] return " ".join(parts).strip()

def _is_empty(v): return v is None or (isinstance(v, str) and v.strip()=="")

def _nonempty_count(vals): return sum(1 for v in vals if not _is_empty(v))

def _row_is_dashes(vals): """ A 'dash row' if every non-empty text cell is only dashes (>=10) from [- – — ], others empty. """ dash_re = r"[-\u2013\u2014]{10,}" saw_long = False for v in vals: if isinstance(v, str) and v.strip(): t = re.sub(r"\s+","", v) if re.fullmatch(dash_re, t): saw_long = True else: return False elif not _is_empty(v): return False return saw_long

def _tight_col_bounds_ws(ws, r1, r2): max_c = ws.max_column nonempty = set() for r in range(r1, r2+1): for c in range(1, max_c+1): if not _is_empty(ws.cell(r,c).value): nonempty.add(c) if not nonempty: return 1, max_c return min(nonempty), max(nonempty)

def _sanitize_sheet_name(name): name = re.sub(r'[:\/?*[]]', " ", str(name)) name = re.sub(r"\s+"," ", name).strip() or "Sheet" return name[:31]

def _unique_sheet_name(base, used): base = _sanitize_sheet_name(base) if base not in used: used.add(base); return base for k in range(2, 1000): alt = f"{base[:31-len(str(k))-3]} ({k})" if alt not in used: used.add(alt); return alt alt = f"Sheet{len(used)+1}"; used.add(alt); return alt

def _copy_block_with_styles(src_ws, dst_ws, r1, c1, r2, c2, dr1=1, dc1=1): # values + styles for r in range(r1, r2+1): for c in range(c1, c2+1): sc = src_ws.cell(r,c) dr = dr1 + (r - r1) dc = dc1 + (c - c1) dc_cell = dst_ws.cell(dr, dc, value=sc.value) dc_cell.number_format = sc.number_format if sc.font: dc_cell.font = copy(sc.font) if sc.fill: dc_cell.fill = copy(sc.fill) if sc.alignment: dc_cell.alignment = copy(sc.alignment) if sc.border: dc_cell.border = copy(sc.border) if sc.protection: dc_cell.protection = copy(sc.protection) # merges for mr in list(src_ws.merged_cells.ranges): if (mr.min_row>=r1 and mr.max_row<=r2 and mr.min_col>=c1 and mr.max_col<=c2): dst_ws.merge_cells( start_row = dr1 + (mr.min_row - r1), start_column = dc1 + (mr.min_col - c1), end_row = dr1 + (mr.max_row - r1), end_column = dc1 + (mr.max_col - c1), ) # column widths for c in range(c1, c2+1): srcL = get_column_letter(c) dstL = get_column_letter(dc1 + (c - c1)) w = src_ws.column_dimensions[srcL].width if w is not None: dst_ws.column_dimensions[dstL].width = w

---------------- Detection ----------------
def _find_banners(ws): """ Detect textual banners per worksheet: dash row (r) + title row (r+1) + dash row (r+2) Return list of tuples: (title_row_1based, title_text_raw) NOTE: we return the MIDDLE row (title), so section can start exactly at the heading. """ max_r, max_c = ws.max_row, ws.max_column grid = [[ws.cell(r,c).value for c in range(1, max_c+1)] for r in range(1, max_r+1)] df = pd.DataFrame(grid) banners = [] r = 0 while r + 2 < len(df): top = df.iloc[r].values mid = df.iloc[r+1].values bot = df.iloc[r+2].values if _row_is_dashes(top) and _row_is_dashes(bot): midtxt_raw = _row_text(mid) midtxt = _norm(midtxt_raw) if midtxt: banners.append((r+2, midtxt_raw)) # (title row 1-based), r+1 (0-based) => +1 for 1-based r += 3 continue r += 1 return banners

def _matches_known_title(text_norm): if not text_norm: return False if any(_norm(t)==text_norm for t in KNOWN_TITLES): return True if re.match(SUMMARY_REGEX, text_norm): return True return False

def _find_title_rows(ws): """ Fallback detection of true titles when no banners exist in the sheet. Return list of (title_row_1based, title_text_raw). A true title row: - has <=2 filled cells, - concatenated normalized text matches known titles or summary regex, - a table row (>=2 filled cells) appears within TABLE_LOOKAHEAD rows after it. """ max_r, max_c = ws.max_row, ws.max_column grid = [[ws.cell(r,c).value for c in range(1, max_c+1)] for r in range(1, max_r+1)] df = pd.DataFrame(grid)

out = []
for i, row in df.iterrows():
    vals = row.values
    filled = _nonempty_count(vals)
    if filled == 0 or filled > 2:
        continue
    text_norm = _norm(_row_text(vals))
    if not _matches_known_title(text_norm):
        continue
    # table soon?
    ok = False
    for j in range(i+1, min(i+1+TABLE_LOOKAHEAD, len(df))):
        if _nonempty_count(df.iloc[j].values) >= 2:
            ok = True
            break
    if ok:
        out.append((i+1, _row_text(vals)))  # 1-based index + raw title
return out
def _dedup_sections(starts): """ starts: list of (row1based, raw_title, norm_title) Deduplicate by normalized title + start proximity (<= DEDUP_PROXIMITY). Keep the earliest start. """ starts = sorted(starts, key=lambda x: x[0]) result = [] for r1, raw, norm in starts: if not result: result.append((r1, raw, norm)) continue prev_r1, prev_raw, prev_norm = result[-1] # same normalized title and starts very close -> treat as one (keep earlier) if norm == prev_norm and abs(r1 - prev_r1) <= DEDUP_PROXIMITY: # merge by keeping earlier; do nothing continue result.append((r1, raw, norm)) return result

---------------- Main processing ----------------
def process_file(input_path): wb_src = load_workbook(input_path, data_only=True) wb_out = Workbook() # ensure at least one sheet exists at the end default_ws = wb_out.active wb_out.remove(default_ws)

used_names = set()
sections_made = 0

for ws in wb_src.worksheets:
    # Collect starts (title rows) from either banners or fallback (not both duplicatively)
    banners = _find_banners(ws)
    if banners:
        starts = [(row, raw, _norm(raw)) for (row, raw) in banners]
    else:
        titles = _find_title_rows(ws)
        starts = [(row, raw, _norm(raw)) for (row, raw) in titles]

    if not starts:
        continue

    # Deduplicate nearby duplicates of the same logical title
    starts = _dedup_sections(starts)

    # Build sections: from each start to the row before next start (or sheet end)
    for i, (start_row, raw_title, norm_title) in enumerate(starts):
        end_row = (starts[i+1][0] - 1) if (i+1 < len(starts)) else ws.max_row
        # Tight column bounds only within this vertical slice
        c1, c2 = _tight_col_bounds_ws(ws, start_row, end_row)
        if c1 > c2 or start_row > end_row:
            continue

        # One SHEET per section: heading + its content
        sheet_name = _unique_sheet_name(raw_title, used_names)
        ws_out = wb_out.create_sheet(title=sheet_name)
        ws_out.sheet_view.rightToLeft = True

        _copy_block_with_styles(ws, ws_out, start_row, c1, end_row, c2, dr1=1, dc1=1)
        sections_made += 1

if sections_made == 0:
    ws_out = wb_out.create_sheet(title="No sections found")
    ws_out.sheet_view.rightToLeft = True
    ws_out["A1"].value = "هیچ سرتیتر/بنری مطابق الگوها پیدا نشد."

base, ext = os.path.splitext(input_path)
out_path = base + "_split.xlsx"
wb_out.save(out_path)
return out_path
---------------- Jupyter: browse & run ----------------
def browse_and_run(): root = Tk(); root.withdraw() file_path = askopenfilename(title="انتخاب فایل اکسل", filetypes=[("Excel files","*.xlsx")]) if not file_path: return None return process_file(file_path)

Run
out_path = browse_and_run() out_path
![Screenshot 2025-11-20 112447](https://github.com/user-attachments/assets/7490c884-00f2-4ec8-9460-c79c53fa188e)
![Screenshot 2025-11-20 112551](https://github.com/user-attachments/assets/e16fd06a-163e-40f5-8c38-03f6d49516b5)
![Screenshot 2025-11-16 233121](https://github.com/user-attachments/assets/c9cae7e6-5118-4a4e-9f6e-004f6d5ba7c2)
