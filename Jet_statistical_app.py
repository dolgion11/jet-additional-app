# app.py
import streamlit as st
import pandas as pd
from pathlib import Path
import build_full_report_pretty as report  # таны үндсэн кодыг дуудаж байна

st.set_page_config(page_title="JET Audit Automation", layout="centered")

st.title("📊 JET Audit Automation Report Generator")

st.markdown("""
Энэхүү апп нь таны оруулсан **GL / TB Excel файлуудаас** 
автомат аудиторын тестийн тайлан үүсгэнэ.
""")

# -----------------------------
# File Upload хэсэг
# -----------------------------
gl_file = st.file_uploader("GL Excel файл оруулна уу", type=["xlsx"])
tb_file = st.file_uploader("TB Excel файл оруулж болно (заавал биш)", type=["xlsx"])

# -----------------------------
# Generate Report Button
# -----------------------------
if st.button("✅ Тайлан үүсгэх"):
    if gl_file:
        # Түр хадгалалт
        gl_path = Path("uploaded_gl.xlsx")
        with open(gl_path, "wb") as f:
         30    f.write(gl_file.read())
31
32    if tb_file:
33        tb_path = Path("uploaded_tb.xlsx")
34        with open(tb_path, "wb") as f:
35            f.write(tb_file.read())
36    else:
37        tb_path = None
38
39    # Initialize variables for report generation
40    report.INPUT_XLSX_GL = gl_path
41    report.INPUT_XLSX_TB = tb_path  # Set TB path if available
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
                file_name="JET_Audit_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("⚠ GL файл заавал оруулах хэрэгтэй.")
