import streamlit as st
import pandas as pd
import build_full_report_pretty as pretty  # өмнө оруулсан функц ашиглана

st.set_page_config(page_title="Тайлангийн генератор", layout="wide")
st.title("📊 Тайлангийн шинээр үүсгэх апп")

uploaded_file = st.file_uploader("Excel файл оруулна уу", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("Файл амжилттай уншигдлаа!")
    pretty.generate_report(df)  # Энэ нь build_full_report_pretty.py дотор байгаа функц байна
