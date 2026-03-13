#!/usr/bin/env python3
"""
Tách file tổng thành 42 file theo phòng ban (cột AJ - col index 35).
Giữ nguyên template gốc (rows 0-38), chỉ copy data rows tương ứng.
"""
import os
import xlrd
from xlrd import XL_CELL_DATE
from xlutils.copy import copy as xl_copy
import xlwt
import datetime

# Lấy file đầu tiên từ thư mục input
input_dir = os.path.join(os.path.dirname(__file__), 'input')
files = [f for f in os.listdir(input_dir) if os.path.isfile(os.path.join(input_dir, f))]
if not files:
    raise ValueError("Không có file nào trong thư mục input")
files.sort()  # Sắp xếp theo thứ tự alphabet
SRC = os.path.join(input_dir, files[0])

OUT_DIR = os.path.join(os.path.dirname(__file__), 'output')
SHEET_NAME = '20_TH_DK_TCT'
DEPT_COL = 35        # AJ (0-based)
HEADER_END = 38      # rows 0-38 are template/header
DATA_START = 39      # data begins at row 39

os.makedirs(OUT_DIR, exist_ok=True)

# --- Read source ---
rb = xlrd.open_workbook(SRC, formatting_info=True)
src_sheet = rb.sheet_by_name(SHEET_NAME)

# Collect data rows grouped by department
dept_rows = {}  # {dept_name: [row_indices]}
for r in range(DATA_START, src_sheet.nrows):
    dept = src_sheet.cell_value(r, DEPT_COL)
    if not dept:
        dept = '_Không có phòng ban'
    dept = str(dept).strip()
    dept_rows.setdefault(dept, []).append(r)

print(f"Tổng phòng ban: {len(dept_rows)}")
print(f"Tổng dòng data: {sum(len(v) for v in dept_rows.values())}")

def safe_filename(name):
    """Tạo tên file an toàn từ tên phòng ban."""
    return name.replace('/', '-').replace('\\', '-').replace(':', '-')

# --- For each department, create a new workbook ---
for dept_name, rows in sorted(dept_rows.items()):
    # Copy the whole workbook (preserves all formatting, merged cells, etc.)
    wb_new = xl_copy(rb)
    ws = wb_new.get_sheet(SHEET_NAME)

    src_si = rb.sheet_names().index(SHEET_NAME)
    src_sh = rb.sheet_by_index(src_si)

    # Build a date style for date cells
    date_style = xlwt.XFStyle()
    date_style.num_format_str = 'DD/MM/YYYY'

    # Write data rows starting after header
    write_row = DATA_START
    for src_r in rows:
        for c in range(src_sh.ncols):
            val = src_sh.cell_value(src_r, c)
            cell_type = src_sh.cell_type(src_r, c)
            if cell_type == XL_CELL_DATE:
                date_tuple = xlrd.xldate_as_tuple(val, rb.datemode)
                dt = datetime.datetime(*date_tuple[:6])
                ws.write(write_row, c, dt, date_style)
            else:
                ws.write(write_row, c, val)
        write_row += 1

    # Clear remaining rows that were in the original but not needed
    for r in range(write_row, src_sh.nrows):
        for c in range(src_sh.ncols):
            ws.write(r, c, '')

    fname = safe_filename(dept_name) + '.xls'
    out_path = os.path.join(OUT_DIR, fname)
    wb_new.save(out_path)
    print(f"  ✓ {fname} ({len(rows)} dòng)")

print(f"\nHoàn tất! {len(dept_rows)} file đã lưu vào: {OUT_DIR}")
