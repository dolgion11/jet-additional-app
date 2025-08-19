# app.py
import streamlit as st
from pathlib import Path
import  build_full_report_pretty as report  # таны үндсэн кодыг дуудаж байна

st.set_page_config(page_title="JET Audit Automation", layout="centered")
st.title("📊 JET Audit Automation Report Generator")

st.markdown(
    """
Энэхүү апп нь таны оруулсан **GL / TB Excel файлуудаас**
автомат аудиторын тестийн тайлан үүсгэнэ.
"""
)

# -----------------------------
# File Upload хэсэг
# -----------------------------
gl_file = st.file_uploader("GL Excel файл оруулна уу", type=["xlsx"])
tb_file = st.file_uploader("TB Excel файл оруулж болно (заавал биш)", type=["xlsx"])

# -----------------------------
# Туслах функц: report модулийн entry-г уян хатан ажиллуулах
# -----------------------------
ENTRY_CANDIDATES = (
    "main",
    "build",
    "run",
    "generate",
    "generate_report",
    "build_report",
    "start",
)

def run_report_module(report_module, gl_path: Path, tb_path: Path | None, out_path: Path):
    """
    1) Модулийн нийтлэг entry нэрүүдийг эрж хайна.
    2) Олдсон callable-г эхлээд (gl, tb, out) параметртэй дуудаж оролдоно.
       Хэрэв параметрийн тоо таарахгүй бол параметргүйгээр дахин дуудна.
    3) Хэрэв callable олдохгүй бол ойлгомжтой алдаа шиднэ.
    """
    # Глобал хувьсагч ашигладаг код байж болох тул урьдчилан онооно
    setattr(report_module, "INPUT_XLSX_GL", gl_path)
    setattr(report_module, "INPUT_XLSX_TB", tb_path if tb_path else gl_path)
    setattr(report_module, "OUTPUT_XLSX", out_path)

    # Entry-г хайх
    entry_func = None
    for name in ENTRY_CANDIDATES:
        func = getattr(report_module, name, None)
        if callable(func):
            entry_func = func
            break

    if entry_func is None:
        raise AttributeError(
            "all_reports_master_merged модульд тохирох entry функц олдсонгүй. "
            "Доорх нэрнүүдийн аль нэгийг экспортлоно уу: "
            + ", ".join(ENTRY_CANDIDATES)
        )

    # Уян хатан дуудлага
    try:
        return entry_func(gl_path, tb_path or gl_path, out_path)
    except TypeError:
        # Хэрэв параметрийн тоо өөр бол параметргүйгээр оролдъё
        return entry_func()

# -----------------------------
# Generate Report Button
# -----------------------------
if st.button("✅ Тайлан үүсгэх"):
    if not gl_file:
        st.error("⚠ GL файл заавал оруулах хэрэгтэй.")
    else:
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

        # Кодоо ажиллуулна
        try:
            with st.spinner("⏳ Тайлан үүсгэж байна..."):
                run_report_module(report, gl_path, tb_path, out_path)

            # Татах линк гаргана
            st.success("✔ Тайлан амжилттай үүсгэлээ!")
            with open(out_path, "rb") as f:
                st.download_button(
                    label="📥 Тайлан татах",
                    data=f,
                    file_name="JET_Audit_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as e:
            st.error("❌ Алдаа гарлаа. Доорх дэлгэрэнгүйг шалгана уу.")
            st.exception(e)
