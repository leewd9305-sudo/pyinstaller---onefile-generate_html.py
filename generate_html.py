import os
import sys
import pandas as pd
import re
from datetime import datetime
import shutil
from tkinter import Tk, filedialog, messagebox
import unicodedata

# ==============================
# 🧭 PyInstaller 리소스 경로 보정
# ==============================
def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


# ==============================
# 📌 엑셀 파일 선택 UI
# ==============================
def select_excel_file():
    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="제작가이드 엑셀 파일을 선택하세요",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not file_path:
        raise FileNotFoundError("엑셀 파일이 선택되지 않았습니다!")

    print(f"\n📌 선택된 엑셀 파일: {file_path}")
    return file_path


# ==============================
# 🛠️ 유틸리티 함수
# ==============================
def sanitize_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", str(name))
    return name.strip()


# ==============================
# 🧩 시트별 HTML → TXT 생성
# ==============================
def generate_html_for_sheet(excel_file_path: str, sheet_name: str, output_dir: str, log_records: list):
    BLOB_BASE_URL = "https://huskb2bstorage.blob.core.windows.net/shopicus/dev_1/guide/03_make/page"
    TOOLTIP_BASE_URL = "https://huskb2bstorage.blob.core.windows.net/shopicus/dev_1/guide/test"

    print(f"\n🚀 [{sheet_name}] 처리 시작")

    folder_name = sheet_name.replace("☆", "").strip()
    sheet_output_dir = os.path.join(output_dir, folder_name)
    os.makedirs(sheet_output_dir, exist_ok=True)

    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=None, dtype=str)
        df = df.fillna("")
    except Exception as e:
        print(f"⚠️ 시트 '{sheet_name}' 로드 실패: {e}")
        return

    # 헤더 2줄 스킵
    df = df.iloc[2:].copy()
    df = df[df[1] != ""].copy()

    # 행 반복 처리
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

        # HTML 템플릿
        html = f"""
        <div style="width:100%; max-width:720px; margin:0 auto; padding:0 16px;
        display:flex; flex-direction:column; align-items:center; gap:30px;
        position:relative; box-sizing:border-box; text-align:center;">
            <div style="background-color:#CCE6FF; border-radius:12px; box-sizing:border-box;
            width:100%; height:fit-content; display:flex; flex-direction:column;
            align-items:center; padding-bottom:30px; position:relative;">
                <img src="{TOOLTIP_BASE_URL}/단색_툴팁.png" alt="단색 제작가이드"
                    style="position:absolute; top:0; left:50%; transform:translateX(-50%);
                    width:130px; height:auto; z-index:10;">
                <h2 style="margin-top:150px; font-size:20px; font-weight:600;">{product_name}</h2>
        """

        for i, file_name in enumerate(image_files, start=1):
            html += f"""
                <div style="margin-top:{55 if i == 1 else 30}px;">
                    <img src="{BLOB_BASE_URL}/{file_name}?ver={i}"
                        alt="{product_name} 이미지 {i}"
                        style="width:100%; max-width:450px;"
                        class="e-rte-image e-imginline">
                </div>
            """

        html += """
            </div>
        </div>
        """

        # 파일 생성
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)

        # 로그 저장
        log_records.append({
            "시트명": sheet_name,
            "순번": seq_str,
            "제품명": product_name,
            "이미지_개수": len(image_files),
            "이미지_파일목록": ", ".join(image_files),
            "HTML_파일경로": output_path
        })

        print(f"✅ [{seq_str}] {product_name} → {output_path}")

    print(f"🎉 [{sheet_name}] 시트 TXT 생성 완료!")


# ==============================
# 🌈 단색+별색 병합
# ==============================
def _extract_mono_content(html_path: str):
    with open(html_path, "r", encoding="utf-8") as f:
        content = f.read()

    product_match = re.search(r'<h2[^>]*>(.*?)</h2>', content)
    product_name = product_match.group(1).strip() if product_match else ""

    image_content_match = re.search(r'</h2\s*>\s*([\s\S]*?)</div>\s*</div>\s*$', content)
    image_content = image_content_match.group(1).strip() if image_content_match else ""

    return product_name, image_content


def _build_combined_block(product_name: str, image_content: str, tooltip_filename: str, tooltip_alt: str, bg_color: str):
    TOOLTIP_BASE_URL = "https://huskb2bstorage.blob.core.windows.net/shopicus/dev_1/guide/test"

    inner_html = f"""
    <h2 style="margin-top:150px; font-size:20px; font-weight:600;">{product_name}</h2>
    {image_content}
    """

    return f"""
    <div style="flex:1; text-align:center; position:relative; overflow:visible;
        display:flex; flex-direction:column; align-items:center;">
        <div style="background-color:{bg_color}; border-radius:12px; box-sizing:border-box;
            width:100%; height:fit-content; display:flex; flex-direction:column;
            align-items:center; padding-bottom:30px; position:relative;">
            <img src="{TOOLTIP_BASE_URL}/{tooltip_filename}" alt="{tooltip_alt}"
                style="position:absolute; top:0; left:50%; transform:translateX(-50%);
                width:130px; height:auto; z-index:10;">
            {inner_html}
        </div>
    </div>
    """


def generate_combined_html(output_dir):
    mono_dir = os.path.join(output_dir, "파일명 리스트(단색)")
    spot_dir = os.path.join(output_dir, "파일명 리스트(별색)")
    combined_dir = os.path.join(output_dir, "combined")
    os.makedirs(combined_dir, exist_ok=True)

    if not os.path.exists(mono_dir) or not os.path.exists(spot_dir):
        print("⚠️ 단색 또는 별색 출력 폴더가 없어 병합을 건너뜁니다.")
        return

    mono_files = sorted(
        [f for f in os.listdir(mono_dir) if f.endswith(".txt")],
        key=lambda x: x.split("_", 1)[0]
    )

    spot_files = {}
    for f in os.listdir(spot_dir):
        if f.endswith(".txt"):
            try:
                product = os.path.splitext(f)[0].split("_", 1)[1]
                spot_files[product] = f
            except:
                continue

    for mono_file in mono_files:
        try:
            seq, product = os.path.splitext(mono_file)[0].split("_", 1)
        except ValueError:
            print(f"⚠️ 단색 파일명 형식 오류: {mono_file}")
            continue

        if product not in spot_files:
            continue

        mono_path = os.path.join(mono_dir, mono_file)
        product_name, image_content = _extract_mono_content(mono_path)

        left_block = _build_combined_block(product_name, image_content, "단색_툴팁.png", "단색 제작가이드", "#CCE6FF")
        right_block = _build_combined_block(product_name, image_content, "별색_툴팁.png", "별색 제작가이드", "#E5F9E0")

        final_html = f"""
        <div style="width:100%; max-width:1420px; margin:0 auto; padding:0 16px;
        display:flex; justify-content:space-between; align-items:flex-start; gap:30px;
        position:relative; box-sizing:border-box; text-align:center;">
            {left_block}
            <div style="position:absolute; top:0; left:50%; transform:translateX(-50%);
            width:1px; height:100%; background-color:#dcdcdc;"></div>
            {right_block}
        </div>
        """

        output_path = os.path.join(combined_dir, f"{seq}_{sanitize_filename(product)}.txt")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(final_html)

        print(f"✨ [{seq}] 병합 완료 → {output_path}")

    print("\n🎉 단색 기준 순서로 병합 TXT 생성 완료!")


# ==============================
# 🚀 전체 실행
# ==============================
if __name__ == "__main__":

    OUTPUT_DIR = resource_path("output")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    try:
        EXCEL_FILE = select_excel_file()
    except FileNotFoundError as e:
        print(f"❌ {e}")
        sys.exit(1)

    LOG_FILE = os.path.join(
        OUTPUT_DIR, f"html_generation_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    )

    try:
        log_records = []

        excel_sheets = pd.ExcelFile(EXCEL_FILE)
        all_sheets = excel_sheets.sheet_names

        # 파일명 리스트 포함된 시트만
        valid_sheets = [s.strip() for s in all_sheets if "파일명 리스트" in s]

        print(f"📄 감지된 시트: {valid_sheets}")

        for sheet in valid_sheets:
            generate_html_for_sheet(EXCEL_FILE, sheet, OUTPUT_DIR, log_records)

        # 전체 로그 저장
        if log_records:
            log_df = pd.DataFrame(log_records)
            log_df.to_csv(LOG_FILE, index=False, encoding="utf-8-sig")
            print(f"\n🧾 전체 로그 저장 완료 → {LOG_FILE}")

            base_time = datetime.now().strftime('%Y%m%d_%H%M%S')

            mono_df = log_df[log_df["시트명"].str.contains("단색")]
            spot_df = log_df[log_df["시트명"].str.contains("별색")]
            normal_df = log_df[
                ~log_df["시트명"].str.contains("단색") &
                ~log_df["시트명"].str.contains("별색")
            ]

            # 단색 로그
            mono_path = os.path.join(OUTPUT_DIR, f"log_mono_{base_time}.csv")
            if not mono_df.empty:
                mono_df.to_csv(mono_path, index=False, encoding="utf-8-sig")
                print(f"🧾 단색 로그 저장 → {mono_path}")

            # 별색 로그
            spot_path = os.path.join(OUTPUT_DIR, f"log_spot_{base_time}.csv")
            if not spot_df.empty:
                spot_df.to_csv(spot_path, index=False, encoding="utf-8-sig")
                print(f"🧾 별색 로그 저장 → {spot_path}")

            # 일반 로그
            normal_path = os.path.join(OUTPUT_DIR, f"log_normal_{base_time}.csv")
            if not normal_df.empty:
                normal_df.to_csv(normal_path, index=False, encoding="utf-8-sig")
                print(f"🧾 일반 로그 저장 → {normal_path}")

        # 병합 실행
        generate_combined_html(OUTPUT_DIR)

        print("\n✨ 모든 TXT 생성 및 병합 완료!")

        # ZIP 압축 생성
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        zip_filename = f"husk_guide_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path_base = os.path.join(downloads_path, zip_filename).replace(".zip", "")

        shutil.make_archive(
            base_name=zip_path_base,
            format="zip",
            root_dir=OUTPUT_DIR
        )

        print(f"\n📦 모든 결과물이 압축되어 저장됨 → {zip_path_base}.zip")

        # ===================================================================
        # 🎉 완료 안내 팝업
        # ===================================================================
        messagebox.showinfo(
            "완료",
            f"제작가이드 변환이 완료되었습니다!\n\n압축 파일 위치:\n{zip_path_base}.zip"
        )

    except Exception as e:
        import traceback
        error_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_log.txt")
        with open(error_log_path, "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        print(f"⚠️ 실행 중 오류 발생! {error_log_path} 파일을 확인하세요.")
