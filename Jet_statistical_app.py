# app.py
import streamlit as st
from pathlib import Path
import runpy
import importlib

st.set_page_config(page_title="JET Audit Automation", layout="centered")
st.title("📊 JET Audit Automation Report Generator")

st.markdown("""
Энэ апп нь таны оруулсан **GL/TB Excel**-ээс `build_full_report_pretty.py`-г ажиллуулж
`final_report.xlsx` тайланг гаргана.
""")

# ---- Файл оруулах
gl_file = st.file_uploader("GL Excel файл (заавал)", type=["xlsx"])
tb_file = st.file_uploader("TB Excel файл (сонголттой)", type=["xlsx"])

# ---- Тайлан үүсгэх
if st.button("✅ Тайлан үүсгэх"):
    if not gl_file:
        st.error("⚠ GL файл заавал оруулна уу.")
        st.stop()

    # Түр хадгалалт
    gl_path = Path("uploaded_gl.xlsx")
    with open(gl_path, "wb") as f:
        f.write(gl_file.read())

    tb_path = None
    if tb_file:
        tb_path = Path("uploaded_tb.xlsx")
        with open(tb_path, "wb") as f:
            f.write(tb_file.read())

    out_path = Path("final_report.xlsx")

    try:
        # Модулийн файлын замыг олно
        report_mod = importlib.import_module("build_full_report_pretty")
        report_file = Path(report_mod.__file__).resolve()

        # Таны модуль глобал хувьсагч ашигладаг тул эндээс утгуудыг өгөөд
        # модулийг __main__ болгон "скриптээр" нь ажиллуулна.
        init_globals = {
            "INPUT_XLSX_GL": gl_path,                 # Path объект байж болох тул шууд Path өгч байна
            "INPUT_XLSX_TB": tb_path or gl_path,      # TB байхгүй бол GL-ийг давхар ашиглана
            "OUTPUT_XLSX":   out_path,                # Гаралтын файл
        }

        with st.spinner("⏳ build_full_report_pretty ажиллаж байна..."):
            runpy.run_path(str(report_file), init_globals=init_globals, run_name="__main__")

        # Амжилттай болсон эсэхийг шалгаад татах товч гаргана
        if out_path.exists():
            st.success("✔ Тайлан амжилттай үүсгэлээ!")
            with open(out_path, "rb") as f:
                st.download_button(
                    label="📥 Тайлан татах",
                    data=f,
                    file_name="JET_Audit_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.error("❌ Тайлангийн файл (final_report.xlsx) үүсээгүй байна. Модулийн логикийг шалгана уу.")

    except Exception as e:
        st.error("❌ Алдаа гарлаа.")
        st.exception(e)
