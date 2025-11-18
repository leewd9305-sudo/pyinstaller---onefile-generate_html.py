import os
import unicodedata
import pandas as pd
import re
from core.config import *
from core.sanitizer import sanitize_filename


# ==============================
# 📄 시트 → HTML/TXT 변환
# ==============================
def generate_html_for_sheet(excel_path, sheet_name, output_dir, log_records):

    sheet_title = sheet_name.replace("☆", "").strip()

    # -----------------------------
    # 🔥 폴더명(original_mode)과 UI 스킨(mode) 분리
    # -----------------------------
    if "단색" in sheet_title:
        original_mode = "단색"
    elif "별색" in sheet_title:
        original_mode = "별색"
    elif "일반" in sheet_title:
        original_mode = "일반"
    else:
        original_mode = "일반"

    # UI 스킨: 일반은 단색처럼 출력
    if original_mode == "단색":
        mode = "단색"
    elif original_mode == "별색":
        mode = "별색"
    elif original_mode == "일반":
        mode = "단색"   # ⭐ 일반 → 단색 스킨 강제

    tooltip_filename = TOOLTIP_MAP[mode]
    tooltip_alt = TOOLTIP_ALT_MAP[mode]
    border_color = COLOR_MAP[mode]

    # -----------------------------
    # 🔥 폴더 구조는 원래 모드 기준 유지
    # -----------------------------
    sheet_output_dir = os.path.join(output_dir, original_mode)
    os.makedirs(sheet_output_dir, exist_ok=True)

    # -----------------------------
    # 이하 기존 HTML 생성 로직 동일
    # -----------------------------
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=str)
    df = df.fillna("")
    df = df.iloc[2:].copy()
    df = df[df[1] != ""].copy()

    for _, row in df.iterrows():

        seq_raw = str(row[1]).strip()
        if not seq_raw:
            continue

        try:
            seq_str = str(int(seq_raw)).zfill(2)
        except:
            seq_str = seq_raw

        product_name = str(row[2]).strip()

        image_files = []
        for i in range(3, len(row)):
            if row[i]:
                clean_val = unicodedata.normalize("NFKC", str(row[i])).strip()
                image_files.append(clean_val)

        if not product_name or not image_files:
            continue

        safe_name = sanitize_filename(product_name)
        output_path = os.path.join(sheet_output_dir, f"{seq_str}_{safe_name}.txt")

        html = f"""
        <div style="width:100%; max-width:720px; margin:0 auto; padding:0 16px;
        display:flex; flex-direction:column; align-items:center; gap:20px;">

            <div style="border:4px solid {border_color}; border-radius:12px; width:100%;
                display:flex; flex-direction:column; align-items:center; padding-bottom:30px; position:relative;">

                <img src="{TOOLTIP_BASE_URL}/{tooltip_filename}"
                    alt="{tooltip_alt}"
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
