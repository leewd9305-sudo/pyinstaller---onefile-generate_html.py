import os
import sys
import pandas as pd
import re
from datetime import datetime
import shutil
from tkinter import Tk, filedialog, messagebox
import unicodedata
from openpyxl import load_workbook   # 🔥 추가


# ==============================
# 🧭 PyInstaller 경로 인식
# ==============================
def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# ==============================
# 📂 엑셀 파일 선택
# ==============================
def select_excel_file():
    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="제작가이드 엑셀 파일 선택",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not file_path:
        raise FileNotFoundError("엑셀 파일이 선택되지 않았습니다!")

    print(f"\n📌 선택된 파일: {file_path}")
    return file_path


# ==============================
# 🔒 파일명 정리
# ==============================
def sanitize_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", str(name))
    return name.strip()


# ==============================
# 📘 엑셀 자동 셀 너비 조정
# ==============================
def save_excel_autowidth(df, path):
    df.to_excel(path, index=False, engine='openpyxl')  # xlsx 파일로 저장

    wb = load_workbook(path)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter

        for cell in col:
            try:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass

        ws.column_dimensions[col_letter].width = (max_length + 2) * 1.2

    wb.save(path)


# ==============================
# 📄 시트 → TXT(HTML) 변환
# ==============================
def generate_html_for_sheet(excel_file_path: str, sheet_name: str, output_dir: str, log_records: list):

    BLOB_BASE_URL = "https://huskb2bstorage.blob.core.windows.net/shopicus/dev_1/guide/03_make/page"
    TOOLTIP_BASE_URL = "https://huskb2bstorage.blob.core.windows.net/shopicus/dev_1/guide/test"

    print(f"\n🚀 [{sheet_name}] 변환 시작")

    folder_name = sheet_name.replace("☆", "").strip()
    sheet_output_dir = os.path.join(output_dir, folder_name)
    os.makedirs(sheet_output_dir, exist_ok=True)

    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=None, dtype=str)
        df = df.fillna("")
    except Exception as e:
        print(f"⚠️ 시트 '{sheet_name}' 로드 실패: {e}")
        return

    df = df.iloc[2:].copy()
    df = df[df[1] != ""].copy()

    for _, row in df.iterrows():

        seq_raw = str(row[1]).strip()
        if not seq_raw:
            continue

        try:
            int(seq_raw)
            seq_str = seq_raw.zfill(2)
        except:
            seq_str = seq_raw

        product_name = str(row[2]).strip()

        image_files = []
        for i in range(3, len(row)):
            val = row[i]
            if not val:
                continue

            clean_val = unicodedata.normalize("NFKC", str(val)).strip()
            image_files.append(clean_val)

        if not product_name or not image_files:
            continue

        safe_name = sanitize_filename(product_name)
        output_path = os.path.join(sheet_output_dir, f"{seq_str}_{safe_name}.txt")

        html = f"""
        <div style="width:100%; max-width:720px; margin:0 auto; padding:0 16px;
        display:flex; flex-direction:column; align-items:center; gap:20px;">

            <div style="border:4px solid #4DA3FF; border-radius:12px; width:100%;
                display:flex; flex-direction:column; align-items:center;
                padding-bottom:30px; position:relative;">

                <img src="{TOOLTIP_BASE_URL}/단색_툴팁.png"
                    alt="단색 제작가이드"
                    style="position:absolute; top:15px; left:50%; transform:translateX(-50%);
                    width:130px; height:auto; z-index:10;">

                <h2 style="margin-top:75px; margin-bottom:30px;
                    font-size:20px; font-weight:600;">{product_name}</h2>
        """

        for i, file_name in enumerate(image_files, start=1):
            html += f"""
                <div style="margin-top:30px;">
                    <img src="{BLOB_BASE_URL}/{file_name}?ver={i}"
                        style="width:100%; max-width:450px;"
                        class="e-rte-image e-imginline">
                </div>
            """

        html += """
            </div>
        </div>
        """

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)

        log_records.append({
            "시트명": sheet_name,
            "순번": seq_str,
            "제품명": product_name,
            "이미지_개수": len(image_files),
            "이미지_파일목록": ", ".join(image_files),
            "HTML_파일경로": output_path
        })

        print(f"✅ [{seq_str}] {product_name} → {output_path}")


# ==============================
# 🔍 단색 TXT → 콘텐츠 추출
# ==============================
def _extract_mono_content(html_path: str):

    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    product_match = re.search(r'<h2[^>]*>(.*?)</h2>', content)
    product_name = product_match.group(1).strip() if product_match else ""

    image_match = re.search(r'</h2[^>]*>([\s\S]*?)</div>\s*</div>\s*$', content)
    image_content = image_match.group(1).strip() if image_match else ""

    return product_name, image_content


# ==============================
# 🧱 단색/별색 공통 블록 생성
# ==============================
def _build_combined_block(product_name, image_content, tooltip_filename, tooltip_alt, border_color):

    TOOLTIP_BASE_URL = "https://huskb2bstorage.blob.core.windows.net/shopicus/dev_1/guide/test"

    return f"""
    <div style="flex:1; text-align:center; display:flex;
        flex-direction:column; align-items:center;">

        <div style="border:4px solid {border_color}; border-radius:12px;
            width:100%; padding-bottom:30px; position:relative;">

            <img src="{TOOLTIP_BASE_URL}/{tooltip_filename}"
                alt="{tooltip_alt}"
                style="position:absolute; top:15px; left:50%; transform:translateX(-50%);
                width:130px; height:auto; z-index:10;">

            <h2 style="margin-top:75px; margin-bottom:30px;
                font-size:20px; font-weight:600;">{product_name}</h2>

            {image_content}

        </div>
    </div>
    """


# ==============================
# 🔗 단색 + 별색 병합 페이지 생성
# ==============================
def generate_combined_html(output_dir):

    mono_dir = os.path.join(output_dir, "파일명 리스트(단색)")
    spot_dir = os.path.join(output_dir, "파일명 리스트(별색)")
    combined_dir = os.path.join(output_dir, "combined")
    os.makedirs(combined_dir, exist_ok=True)

    if not os.path.exists(mono_dir) or not os.path.exists(spot_dir):
        print("⚠️ 병합 불가 — 단색/별색 폴더 없음")
        return

    mono_files = sorted(
        [f for f in os.listdir(mono_dir) if f.endswith(".txt")],
        key=lambda x: x.split("_", 1)[0]
    )

    spot_files = {
        os.path.splitext(f)[0].split("_", 1)[1]: f
        for f in os.listdir(spot_dir)
        if f.endswith(".txt")
    }

    for mono_file in mono_files:

        try:
            seq, product = os.path.splitext(mono_file)[0].split("_", 1)
        except:
            continue

        if product not in spot_files:
            continue

        mono_path = os.path.join(mono_dir, mono_file)
        product_name, image_content = _extract_mono_content(mono_path)

        left_block = _build_combined_block(
            product_name, image_content,
            "단색_툴팁.png", "단색 제작가이드", "#4DA3FF"
        )

        right_block = _build_combined_block(
            product_name, image_content,
            "별색_툴팁.png", "별색 제작가이드", "#24CF7F"
        )

        final_html = f"""
        <div style="width:100%; max-width:1420px; margin:0 auto; padding:0 16px;
        display:flex; justify-content:space-between; gap:30px; position:relative;">

            {left_block}

            <div style="position:absolute; top:0; left:50%; transform:translateX(-50%);
                width:1px; height:100%; background:#dcdcdc;"></div>

            {right_block}
        </div>
        """

        output_path = os.path.join(combined_dir, f"{seq}_{sanitize_filename(product)}.txt")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(final_html)

        print(f"✨ 병합 완료 → {output_path}")

    print("🎉 병합 TXT 생성 완료!")


# ==============================
# 🏁 메인 실행부
# ==============================
if __name__ == "__main__":

    OUTPUT_DIR = resource_path("output")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    try:
        EXCEL_FILE = select_excel_file()
    except:
        sys.exit(1)

    LOG_TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')

    try:

        log_records = []

        excel_sheets = pd.ExcelFile(EXCEL_FILE)
        all_sheets = excel_sheets.sheet_names

        valid_sheets = [s.strip() for s in all_sheets if "파일명 리스트" in s]

        for sheet in valid_sheets:
            generate_html_for_sheet(EXCEL_FILE, sheet, OUTPUT_DIR, log_records)

        # 로그 생성
        if log_records:
            log_df = pd.DataFrame(log_records)

            # 저장 경로를 xlsx로 변경
            LOG_XLSX = os.path.join(OUTPUT_DIR, f"html_log_{LOG_TIMESTAMP}.xlsx")
            save_excel_autowidth(log_df, LOG_XLSX)

            mono_df = log_df[log_df["시트명"].str.contains("단색", na=False)]
            spot_df = log_df[log_df["시트명"].str.contains("별색", na=False)]
            normal_df = log_df[
                ~log_df["시트명"].str.contains("단색", na=False) &
                ~log_df["시트명"].str.contains("별색", na=False)
            ]

            if not mono_df.empty:
                save_excel_autowidth(mono_df, os.path.join(OUTPUT_DIR, f"log_mono_{LOG_TIMESTAMP}.xlsx"))

            if not spot_df.empty:
                save_excel_autowidth(spot_df, os.path.join(OUTPUT_DIR, f"log_spot_{LOG_TIMESTAMP}.xlsx"))

            if not normal_df.empty:
                save_excel_autowidth(normal_df, os.path.join(OUTPUT_DIR, f"log_normal_{LOG_TIMESTAMP}.xlsx"))

        generate_combined_html(OUTPUT_DIR)

        # ZIP 압축 생성
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        zip_filename = f"husk_guide_output_{LOG_TIMESTAMP}.zip"
        zip_path_base = os.path.join(downloads_path, zip_filename).replace(".zip", "")

        shutil.make_archive(
            base_name=zip_path_base,
            format="zip",
            root_dir=OUTPUT_DIR
        )

        messagebox.showinfo(
            "완료",
            f"제작가이드 변환이 완료되었습니다!\n압축 파일 위치:\n{zip_path_base}.zip"
        )

    except Exception:
        import traceback
        error_path = os.path.join(os.path.dirname(__file__), "error_log.txt")
        with open(error_path, "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        print(f"⚠️ 오류 발생 → {error_path}")
