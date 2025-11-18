import os
import re
from core.sanitizer import sanitize_filename
from core.config import TOOLTIP_BASE_URL


def _extract_content(path):
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    product_match = re.search(r'<h2[^>]*>(.*?)</h2>', content)
    product_name = product_match.group(1).strip() if product_match else ""

    image_match = re.search(r'</h2[^>]*>([\s\S]*?)</div>\s*</div>\s*$', content)
    image_content = image_match.group(1).strip() if image_match else ""

    return product_name, image_content


def _build_block(product_name, content, tooltip_filename, tooltip_alt, color):

    return f"""
    <div style="flex:1; text-align:center; display:flex; flex-direction:column; align-items:center;">
        <div style="border:4px solid {color}; border-radius:12px; width:100%; padding-bottom:30px; position:relative;">

            <img src="{TOOLTIP_BASE_URL}/{tooltip_filename}"
                alt="{tooltip_alt}"
                style="position:absolute; top:15px; left:50%; transform:translateX(-50%);
                width:130px; height:auto; z-index:10;">

            <h2 style="margin-top:75px; margin-bottom:30px;
                font-size:20px; font-weight:600;">{product_name}</h2>

            {content}
        </div>
    </div>
    """


def generate_combined_html(output_dir):

    mono_dir = os.path.join(output_dir, "단색")
    spot_dir = os.path.join(output_dir, "별색")
    combined_dir = os.path.join(output_dir, "combined")
    os.makedirs(combined_dir, exist_ok=True)

    if not os.path.exists(mono_dir) or not os.path.exists(spot_dir):
        print("⚠️ 병합 불가 — '단색' 또는 '별색' 폴더 없음")
        return

    mono_files = sorted(
        [f for f in os.listdir(mono_dir) if f.endswith(".txt")],
        key=lambda x: x.split("_", 1)[0]
    )

    spot_map = {
        os.path.splitext(f)[0].split("_", 1)[1]: f
        for f in os.listdir(spot_dir)
        if f.endswith(".txt")
    }

    for mono_file in mono_files:

        try:
            seq, product = os.path.splitext(mono_file)[0].split("_", 1)
        except:
            continue

        if product not in spot_map:
            continue

        mono_path = os.path.join(mono_dir, mono_file)
        spot_path = os.path.join(spot_dir, spot_map[product])

        mono_name, mono_content = _extract_content(mono_path)
        spot_name, spot_content = _extract_content(spot_path)

        left_block = _build_block(mono_name, mono_content, "단색_툴팁.png", "단색 제작가이드", "#4DA3FF")
        right_block = _build_block(spot_name, spot_content, "별색_툴팁.png", "별색 제작가이드", "#24CF7F")

        html = f"""
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
            f.write(html)

    print("🎉 병합 완료!")
