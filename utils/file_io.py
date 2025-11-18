import os
import shutil

def zip_output_only(output_dir, timestamp):
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    zip_name = f"husk_guide_output_{timestamp}"
    zip_base = os.path.join(downloads, zip_name)

    # 🔥 output 폴더만 압축
    shutil.make_archive(
        base_name=zip_base,
        format="zip",
        root_dir=output_dir
    )

    return f"{zip_base}.zip"
