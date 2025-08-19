
# app.py
import streamlit as st
import pandas as pd
from pathlib import Path
import build_full_report_pretty_with_main as report  # шинэчлэгдсэн кодын нэр

st.set_page_config(page_title="JET Audit Automation", layout="centered")

st.title("📊 JET Audit Automation Report Generator")

st.markdown("""
Энэхүү апп нь таны оруулсан **GL / TB Excel файлуудаас** 
автомат аудиторын тестийн тайлан үүсгэнэ.
""")

# -----------------------------
# File Upload хэсэг
# -----------------------------
gl_file = st.file_uploader("Excel файл оруулна уу", type=["xlsx"])

# -----------------------------
# Generate Report Button
# -----------------------------
if st.button("✅ Тайлан үүсгэх"):
    if gl_file:
        # Түр хадгалалт
        gl_path = Path("uploaded_data.xlsx")
        with open(gl_path, "wb") as f:
            f.write(gl_file.read())

        # Тайлангийн гаралт
        output_path = Path("final_report.xlsx")

        # Кодоо ажиллуулна
        with st.spinner("⏳ Тайлан үүсгэж байна..."):
            report.main(gl_path=gl_path, output_path=output_path)

        # Татах линк гаргана
        st.success("✔ Тайлан амжилттай үүсгэлээ!")
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Тайлан татах",
                data=f,
                file_name="JET_Audit_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("⚠ Excel файл заавал оруулах хэрэгтэй.")
