import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from core.config import YELLOW_HEX


# ==============================
# 🔍 노란색 행 판별
# ==============================
def find_changed_rows(excel_path, sheet_name):
    wb = load_workbook(excel_path, data_only=True)

    if sheet_name not in wb.sheetnames:
        print(f"⚠️ '{sheet_name}' 시트를 찾을 수 없습니다.")
        return set()

    ws = wb[sheet_name]
    changed_rows = set()

    for row in ws.iter_rows(min_row=3):
        for cell in row:
            fill = cell.fill
            if fill and fill.start_color and fill.start_color.rgb:
                rgb = fill.start_color.rgb.upper()
                if rgb.endswith("FFFF00"):
                    changed_rows.add(cell.row)
                    break

    return changed_rows


# ==============================
# 🟡 변경된 row 하이라이트 + 자동 컬럼 너비
# ==============================
def save_excel_with_highlight(df, path, changed_rows):
    df.to_excel(path, index=False, engine='openpyxl')

    wb = load_workbook(path)
    ws = wb.active

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 자동 너비 조정
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter

        for cell in col:
            try:
                max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        ws.column_dimensions[col_letter].width = (max_length + 2) * 1.2

    # 행 강조
    for src_row in changed_rows:
        log_row = src_row - 1
        if log_row >= 2:
            for cell in ws[log_row]:
                cell.fill = yellow_fill

    wb.save(path)
