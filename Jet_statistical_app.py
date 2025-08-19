# app.py
import streamlit as st
import pandas as pd
from pathlib import Path
import build_full_report_pretty as report  # таны үндсэн кодыг дуудаж байна

st.set_page_config(page_title="JET statistics Automation", layout="centered")

st.title("📊 JET statistics Automation Report Generator")

st.markdown("""
Энэхүү апп нь таны оруулсан **GL / TB Excel файлуудаас** 
автомат аудиторын тестийн тайлан үүсгэнэ.
""")

# -----------------------------
# File Upload хэсэг
# -----------------------------
gl_file = st.file_uploader(" Excel файл оруулна уу", type=["xlsx"])


# -----------------------------
# Generate Report Button
# -----------------------------
if st.button("✅ Тайлан үүсгэх"):
    if gl_file:
        # Түр хадгалалт
        gl_path = Path("uploaded_gl.xlsx")
        with open(gl_path, "wb") as f:
            f.write(gl_file.read())

        if tb_file:
            tb_path = Path("uploaded_tb.xlsx")
            with open(tb_path, "wb") as f:
                f.write(tb_file.read())
        else:
            tb_path = None

        # Таны кодын замуудыг өөрчилнө
        report.INPUT_XLSX_GL = gl_path
        report.INPUT_XLSX_TB = tb_path if tb_file else gl_path
        report.OUTPUT_XLSX   = Path("final_report.xlsx")

        # Кодоо ажиллуулна
        with st.spinner("⏳ Тайлан үүсгэж байна..."):
            report.main()

        # Татах линк гаргана
        st.success("✔ Тайлан амжилттай үүсгэлээ!")
        with open("final_report.xlsx", "rb") as f:
            st.download_button(
                label="📥 Тайлан татах",
                data=f,
                file_name="JET_statistics_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("⚠ GL файл заавал оруулах хэрэгтэй.")
