import os
import pandas as pd
from datetime import datetime

from utils.dialogs import select_excel_file, show_info, show_error
from utils.path_helper import resource_path
from utils.file_io import zip_output_only

from core.excel_reader import find_changed_rows, save_excel_with_highlight
from core.html_generator import generate_html_for_sheet
from core.merger import generate_combined_html


if __name__ == "__main__":

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    OUTPUT_DIR = os.path.join(BASE_DIR, "output")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    try:
        EXCEL_FILE = select_excel_file()
    except:
        exit(1)

    LOG_TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')

    changed_rows_map = {}
    log_records = []

    try:
        excel_sheets = pd.ExcelFile(EXCEL_FILE)
        valid_sheets = [
            s for s in excel_sheets.sheet_names
            if any(k in s for k in ["단색", "별색", "일반"])
        ]

        # ============================
        # 1) 시트별 HTML 생성
        # ============================
        for sheet in valid_sheets:
            changed_rows_map[sheet] = find_changed_rows(EXCEL_FILE, sheet)
            generate_html_for_sheet(EXCEL_FILE, sheet, OUTPUT_DIR, log_records)

        # ============================
        # 2) 로그 생성 (원본 동일)
        # ============================
        if log_records:
            log_df = pd.DataFrame(log_records)

            # 전체 로그 저장
            all_log_path = os.path.join(
                OUTPUT_DIR, f"html_log_all_{LOG_TIMESTAMP}.xlsx"
            )
            log_df.to_excel(all_log_path, index=False, engine="openpyxl")

            mono_df = log_df[log_df["시트명"].str.contains("단색", na=False)]
            spot_df = log_df[log_df["시트명"].str.contains("별색", na=False)]
            normal_df = log_df[log_df["시트명"].str.contains("일반", na=False)]

            mono_sheet_name = next((s for s in valid_sheets if "단색" in s), None)
            spot_sheet_name = next((s for s in valid_sheets if "별색" in s), None)
            normal_sheet_name = next((s for s in valid_sheets if "일반" in s), None)

            # 단색 로그
            if not mono_df.empty and mono_sheet_name:
                save_excel_with_highlight(
                    mono_df,
                    os.path.join(OUTPUT_DIR, f"log_mono_{LOG_TIMESTAMP}.xlsx"),
                    changed_rows_map.get(mono_sheet_name, set())
                )

            # 별색 로그
            if not spot_df.empty and spot_sheet_name:
                save_excel_with_highlight(
                    spot_df,
                    os.path.join(OUTPUT_DIR, f"log_spot_{LOG_TIMESTAMP}.xlsx"),
                    changed_rows_map.get(spot_sheet_name, set())
                )

            # 일반 로그
            if not normal_df.empty and normal_sheet_name:
                save_excel_with_highlight(
                    normal_df,
                    os.path.join(OUTPUT_DIR, f"log_normal_{LOG_TIMESTAMP}.xlsx"),
                    changed_rows_map.get(normal_sheet_name, set())
                )

        # ============================
        # 3) 단색 + 별색 병합 페이지
        # ============================
        generate_combined_html(OUTPUT_DIR)

        # ============================
        # 4) output 결과물만 zip 압축 (원함)
        # ============================
        zip_path = zip_output_only(OUTPUT_DIR, LOG_TIMESTAMP)

        show_info(f"제작가이드 변환 완료!\nZIP 파일 위치:\n{zip_path}")

    except Exception as e:
        show_error(str(e))
        raise
