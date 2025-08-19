# app.py
import streamlit as st
from pathlib import Path
import build_full_report_pretty as report  # таны үндсэн код

st.set_page_config(page_title="JET statistics Automation", layout="centered")
st.title("📊 JET statistics Automation Report Generator")

st.markdown(
    "Энэхүү апп нь таны оруулсан **Excel** файлын олон sheet-үүдээс тайлан үүсгэнэ."
)

# ---- 1. Файл оруулах (зөвхөн 1 файл) ----
uploaded = st.file_uploader("Excel файл оруулна уу", type=["xlsx"])

# ---- 2. Тайлан үүсгэх ----
if st.button("✅ Тайлан үүсгэх"):
    if not uploaded:
        st.error("⚠ Excel файл заавал оруулна уу.")
        st.stop()

    # Түр хадгалалт
    gl_path = Path("uploaded.xlsx")
    with open(gl_path, "wb") as f:
        f.write(uploaded.getbuffer())

    # Танай тайлангийн модулийн параметрүүд
    report.INPUT_XLSX_GL = gl_path
    report.INPUT_XLSX_TB = gl_path        # TB тусдаа байхгүй тул ижил файлыг зааж өгнө
    report.OUTPUT_XLSX   = Path("final_report.xlsx")

    try:
        with st.spinner("⏳ Тайлан үүсгэж байна..."):
            report.main()
    except Exception as e:
        st.exception(e)
    else:
        st.success("✔ Тайлан амжилттай үүсгэлээ!")
        with open(report.OUTPUT_XLSX, "rb") as f:
            st.download_button(
                label="📥 Тайлан татах",
                data=f,
                file_name="JET_statistics_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
