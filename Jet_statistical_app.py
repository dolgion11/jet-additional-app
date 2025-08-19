# app.py
import streamlit as st
from pathlib import Path
import pandas as pd
import build_full_report_pretty as report
import tempfile
import os

# App configuration
st.set_page_config(page_title="JET Audit Automation", layout="centered")
st.title("📊 JET Audit Automation Report Generator")

st.markdown("""
Энэхүү апп нь таны оруулсан **GL/TB Excel файлуудаас** аудитын тайлан автоматаар үүсгэнэ.
""")

# Custom CSS for better appearance
st.markdown("""
<style>
    .stDownloadButton button {
        width: 100%;
    }
    .stFileUploader {
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# File upload section
with st.expander("📁 Файл оруулах", expanded=True):
    gl_file = st.file_uploader("GL Excel файл оруулна уу (заавал)", type=["xlsx"])
    tb_file = st.file_uploader("TB Excel файл оруулна уу (заавал биш)", type=["xlsx"])
    
    # Materiality inputs
    st.markdown("**Materiality тохиргоо**")
    col1, col2 = st.columns(2)
    with col1:
        ctt = st.number_input("Threshold (CTT)", value=135050000, step=1000000)
    with col2:
        pm = st.number_input("Performance Materiality (PM)", value=1620600000, step=1000000)

# Generate report button
if st.button("✅ Тайлан үүсгэх", use_container_width=True):
    if not gl_file:
        st.error("⚠ GL файл заавал оруулах шаардлагатай!")
        st.stop()
    
    with st.spinner("⏳ Тайлан үүсгэж байна..."):
        try:
            # Create temporary directory
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                
                # Save uploaded files
                gl_path = temp_path / "uploaded_gl.xlsx"
                with open(gl_path, "wb") as f:
                    f.write(gl_file.getvalue())
                
                tb_path = None
                if tb_file:
                    tb_path = temp_path / "uploaded_tb.xlsx"
                    with open(tb_path, "wb") as f:
                        f.write(tb_file.getvalue())
                
                # Configure report settings
                report.SRC_PATH = str(gl_path)
                report.TB_SHEET = "TB"
                report.GL_SHEET = "GL"
                report.MAT_SHEET = "Materiality"
                report.CTT = ctt
                report.PM = pm
                
                # Set output path
                output_path = temp_path / "final_report.xlsx"
                report.OUT_PATH = str(output_path)
                
                # Generate the report
                report.build_reconciliation = report.build_reconciliation
                report.build_materiality = report.build_materiality
                report.build_je_by_account_like_pivot = report.build_je_by_account_like_pivot
                report.build_by_month = report.build_by_month
                report.build_day_group = report.build_day_group
                report.build_by_user = report.build_by_user
                report.build_by_dow = report.build_by_dow
                report.build_net_to_zero = report.build_net_to_zero
                
                # Execute the report generation
                report.build_reconciliation = report.build_reconciliation
                report.build_materiality = report.build_materiality
                report.build_je_by_account_like_pivot = report.build_je_by_account_like_pivot
                report.build_by_month = report.build_by_month
                report.build_day_group = report.build_day_group
                report.build_by_user = report.build_by_user
                report.build_by_dow = report.build_by_dow
                report.build_net_to_zero = report.build_net_to_zero
                
                # Build data
                TB = pd.read_excel(report.SRC_PATH, sheet_name=report.TB_SHEET) if report.TB_SHEET in pd.ExcelFile(report.SRC_PATH).sheet_names else pd.DataFrame()
                GL = pd.read_excel(report.SRC_PATH, sheet_name=report.GL_SHEET) if report.GL_SHEET in pd.ExcelFile(report.SRC_PATH).sheet_names else pd.DataFrame()
                MAT_RAW = pd.read_excel(report.SRC_PATH, sheet_name=report.MAT_SHEET) if report.MAT_SHEET in pd.ExcelFile(report.SRC_PATH).sheet_names else pd.DataFrame()
                
                recon = report.build_reconciliation(TB, GL)
                materiality_df = report.build_materiality(GL, report.CTT, report.PM)
                pivot, total_entries, total_value, max_entries, min_entries, most_list, least_list = report.build_je_by_account_like_pivot(GL)
                by_month = report.build_by_month(GL)
                by_day_group = report.build_day_group(GL)
                by_user = report.build_by_user(GL)
                by_dow = report.build_by_dow(GL)
                net_zero = report.build_net_to_zero(GL)
                
                # Write to Excel
                wb = report.Workbook()
                
                # 1) Reconcilation
                ws = wb.active; ws.title = "Reconcilation"
                report.write_table(ws, recon,
                    money_cols={"Opening Balance per TB","Ending balance per TB (Total Correct)",
                                "Movement per TB 1","Movement per GL","Difference Rounded"})
                
                # 2) Materiality
                ws = wb.create_sheet("Materiality")
                report.write_table(
                    ws, materiality_df,
                    start_row=10, start_col=2,
                    money_cols={"Total Amount (in MNT)"},
                    int_cols={"Number of Line Items Involved"},
                    percent_cols={"Percentage","Amount Percentage"},
                    freeze=True, show_grid=False, title_text="Testing:"
                )
                
                # 3) JE_by_Account
                ws = wb.create_sheet("JE_by_Account")
                ws["B9"] = "Testing:"; ws["B9"].font = report.Font(bold=True, color="CC0000")
                pivot_cols = ["Account Number","Account Name","Total Number of Entries","Value of transactions"]
                end_r = report.write_table(ws, pivot[pivot_cols], start_row=10, start_col=2,
                                        money_cols={"Value of transactions"},
                                        int_cols={"Total Number of Entries"})
                tot_r = end_r + 1
                ws.cell(tot_r, 4, "Total").font=report.BOLD; ws.cell(tot_r, 4).alignment=report.Alignment(horizontal="right")
                ws.cell(tot_r, 5, total_entries).number_format='0'; ws.cell(tot_r,5).border=report.BORDER
                ws.cell(tot_r, 6, total_value).number_format='#,##0'; ws.cell(tot_r,6).border=report.BORDER
                ws.cell(tot_r+2, 2, "Comments:").font=report.BOLD
                hdr_r = tot_r + 4
                for j, txt in enumerate(["Number of entries","Number of accounts","Account name","Amount"], start=3):
                    c = ws.cell(hdr_r, j, txt); c.font=report.BOLD; c.fill=report.HDR_FILL
                    c.alignment=report.Alignment(horizontal="center"); c.border=report.BORDER
                row = hdr_r + 2
                ws.cell(row, 2, "Most used account").font=report.BOLD
                ws.cell(row, 3, max_entries).number_format='0'; ws.cell(row,3).border=report.BORDER
                ws.cell(row, 4, len(most_list)).number_format='0'; ws.cell(row,4).border=report.BORDER
                for i, rec in most_list.iterrows():
                    r = row + i
                    ws.cell(r, 5, rec["Account Name"]).border=report.BORDER
                    cc = ws.cell(r, 6, float(rec["Amount"])); cc.number_format='#,##0'; cc.border=report.BORDER
                row = hdr_r + 2 + max(1, len(most_list)) + 3
                ws.cell(row, 2, "Least used accounts").font=report.BOLD
                ws.cell(row, 3, min_entries).number_format='0'; ws.cell(row,3).border=report.BORDER
                ws.cell(row, 4, len(least_list)).number_format='0'; ws.cell(row,4).border=report.BORDER
                for i, rec in least_list.iterrows():
                    r = row + i
                    ws.cell(r, 5, rec["Account Name"]).border=report.BORDER
                    cc = ws.cell(r, 6, float(rec["Amount"])); cc.number_format='#,##0'; cc.border=report.BORDER
                ws.freeze_panes = "B11"; ws.sheet_view.showGridLines = False
                
                # 4-8) Other sheets
                ws = wb.create_sheet("JE_by_Month")
                report.write_table(ws, by_month, money_cols={"Total Amount (in MNT)"},
                                int_cols={"Total Number of Line Items"})
                
                ws = wb.create_sheet("JE_by_Day_of_Month")
                report.write_table(ws, by_day_group, money_cols={"Total Amount (in MNT)"},
                                int_cols={"Total Number of Line Items"})
                
                ws = wb.create_sheet("JE Distribution by User")
                report.write_table(ws, by_user, money_cols={"Total Amount (in MNT)"},
                                int_cols={"Total Number of Line Items"})
                
                ws = wb.create_sheet("JE_by_Day_of_Week")
                report.write_table(ws, by_dow, int_cols={"Total Number of Line Items","Days in the year"})
                
                ws = wb.create_sheet("Net_to_Zero_Test")
                report.write_table(ws, net_zero, money_cols={"Sum of Transaction"})
                
                # RAW sheets
                ws = wb.create_sheet("TB");  report.write_table(ws, TB,  freeze=False, show_grid=True)
                ws = wb.create_sheet("GL");  report.write_table(ws, GL,  freeze=False, show_grid=True)
                ws = wb.create_sheet("materiality_raw"); report.write_table(ws, MAT_RAW, freeze=False, show_grid=True)
                
                wb.save(output_path)
                
                # Show success message
                st.success("🎉 Тайлан амжилттай боловсрууллаа!")
                
                # Download button
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Тайлан татах",
                        data=f,
                        file_name="JET_Audit_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
        except Exception as e:
            st.error(f"Тайлан үүсгэхэд алдаа гарлаа: {str(e)}")
            st.error("Алдааны дэлгэрэнгүй:")
            st.exception(e)
