
# build_full_report_pretty_with_main.py
import pandas as pd

def main(gl_path, output_path):
    # Бүх sheet-үүдийг уншина
    all_sheets = pd.read_excel(gl_path, sheet_name=None)

    # Жишээ болгож эхний sheet-г буцаана
    sheet_names = list(all_sheets.keys())

    # Зүгээр л бүх sheet-ийг нэг файлд хуулаад хадгалах (жишээ код)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
