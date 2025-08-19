# app.py
import streamlit as st
import pandas as pd
from pathlib import Path
import build_full_report_pretty as report

# Тохиргоо
st.set_page_config(page_title="JET Audit Automation", layout="centered")
st.title("📊 JET Audit Automation Report Generator")

st.markdown("""
Энэхүү апп нь таны оруулсан **GL / TB Excel файлуудаас** 
автомат аудитын тайлан үүсгэнэ.
""")

# -----------------------------
# Файл ачааллын хэсэг
# -----------------------------
gl_file = st.file_uploader("GL Excel файл оруулна уу (заавал)", type=["xlsx"])
tb_file = st.file_uploader("TB Excel файл оруулна уу (заавал биш)", type=["xlsx"])

# -----------------------------
# Тайлан үүсгэх товч
# -----------------------------
if st.button("✅ Тайлан үүсгэх"):
    if not gl_file:
        st.error("⚠ GL файл заавал оруулах шаардлагатай!")
        st.stop()
    
    try:
        # GL файлыг түр хадгалах
        gl_path = Path("uploaded_gl.xlsx")
        with open(gl_path, "wb") as f:
            f.write(gl_file.getvalue())  # getvalue() ашиглах нь илүү найдвартай
            
        # TB файл байгаа эсэхийг шалгах
        tb_path = None
        if tb_file:
            tb_path = Path("uploaded_tb.xlsx")
            with open(tb_path, "wb") as f:
                f.write(tb_file.getvalue())
        
        # Тайлангийн тохиргоо
        report.INPUT_XLSX_GL = str(gl_path)  # Path объектыг string болгож өгөх
        report.INPUT_XLSX_TB = str(tb_path) if tb_path else str(gl_path)
        report.OUTPUT_XLSX = str(Path("final_report.xlsx"))
        
        # Тайлан үүсгэх
        with st.spinner("⏳ Тайлан үүсгэж байна..."):
            report.main()  # report модулийн main() функц дуудагдах ёстой
            
        # Үр дүнг харуулах
        st.success("🎉 Тайлан амжилттай боловсрууллаа!")
        
        # Татах холбоос
        if Path("final_report.xlsx").exists():
            with open("final_report.xlsx", "rb") as f:
                st.download_button(
                    label="📥 Тайлан татах",
                    data=f,
                    file_name="JET_Audit_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Тайлангийн файл үүсээгүй байна. Алдааны мэдээллийг шалгана уу.")
            
    except Exception as e:
        st.error(f"Алдаа гарлаа: {str(e)}")
        st.error("Тайлан үүсгэхэд алдаа гарлаа. Дахин оролдоно уу.")
