# all_reports_master.py
# -*- coding: utf-8 -*-
# pip install pandas openpyxl xlsxwriter

import re
from pathlib import Path
import pandas as pd
import numpy as np

# ---------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------
INPUT_XLSX_GL = Path(r"C:\pYTHON\steppe road data.xlsx")  # GL
INPUT_XLSX_TB = Path(r"C:\pYTHON\Steppe Road of Development LLC-JET Statistics 1.xlsx")  # TB (trial balance)
OUTPUT_XLSX   = Path(r"C:\pYTHON\All_Tests_Report1.xlsx")
TITLE_DATE    = "31 December 2024"

# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------
ALIASES = {
    "Данс": ["Данс","Account","GL Account","Account No","Account Number"],
    "Дансны нэр": ["Дансны нэр","Account Name","GL Account Name","Name"],
    "Огноо": ["Огноо","Posted Date","Posting Date","Date","GL Date","Үүсгэсэн огноо"],
    "Гүйлгээний дугаар": ["Гүйлгээний дугаар","Document No","Voucher No","Entry No","Trans No"],
    "Харьцсан дансны нэр": ["Харьцсан дансны нэр","Counter Account Name","Offset Account Name"],
    "Харьцсан данс": ["Харьцсан данс","Counter Account","Offset Account"],
    "Баримт дугаар": ["Баримтын дугаар","Invoice No","Bill No","Receipt No"],
    "Валют": ["Валют","Currency","Currency Code"],
    "Ханш": ["Ханш","Exchange Rate","Rate"],
    "Валютын дүн": ["Валютын дүн","Foreign Amount","FCY Amount","Amount (FCY)"],
    "Дебет дүн": ["Дебет дүн","Debit","Debit Amount","Debit (MNT)"],
    "Кредит дүн": ["Кредит дүн","Credit","Credit Amount","Credit (MNT)"],
    "Transaction": ["Transaction","Amount","Transaction Amount","Гүйлгээний дүн"],
    "Гүйлгээний утга": ["Гүйлгээний утга","Description","Memo","Narration"],
    "ABS": ["ABS","Absolute"],
    "Бүртгэсэн хэрэглэгч": ["Бүртгэсэн хэрэглэгч","User","Posted By","Created By"],
    "Type": ["Type","Төрөл"],
    "Day of the week": ["Day of the week"],
    "Day of the week /posted date/": ["Day of the week /posted date/"],
    "Цонхны нэр": ["Цонхны нэр","Window Name"],
}

def clean_sheet_name(name: str) -> str:
    for ch in '[]:*?/\\':
        name = name.replace(ch, '-')
    return name[:31]  # Excel limit

def load_first_sheet(xlsx_path: Path, prefer_keywords: list[str]) -> tuple[pd.DataFrame, str]:
    xls = pd.ExcelFile(xlsx_path)
    # exact match
    for kw in prefer_keywords:
        exact = next((s for s in xls.sheet_names if s.strip().lower() == kw.lower()), None)
        if exact:
            df = pd.read_excel(xls, sheet_name=exact, engine="openpyxl")
            return df.rename(columns={c: str(c).strip() for c in df.columns}), exact
    # contains
    for kw in prefer_keywords:
        contains = next((s for s in xls.sheet_names if kw.lower() in s.strip().lower()), None)
        if contains:
            df = pd.read_excel(xls, sheet_name=contains, engine="openpyxl")
            return df.rename(columns={c: str(c).strip() for c in df.columns}), contains
    # fallback first
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], engine="openpyxl")
    return df.rename(columns={c: str(c).strip() for c in df.columns}), xls.sheet_names[0]

def load_gl(path: Path) -> tuple[pd.DataFrame, str]:
    return load_first_sheet(path, ["gl"])

def load_tb(path: Path) -> tuple[pd.DataFrame, str]:
    return load_first_sheet(path, ["tb", "trial balance", "balance", "jet", "statistics"])

def to_number(x):
    if pd.isna(x): return pd.NA
    s = str(x).strip()
    neg = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^\d\.\-]", "", s)
    try:
        v = float(s) if s else pd.NA
        if v is pd.NA: return pd.NA
        return -v if neg else v
    except:
        return pd.NA

def match_col(target, available):
    # exact
    for c in available:
        if c.strip().lower() == target.strip().lower(): return c
    # alias exact
    for alt in ALIASES.get(target, []):
        for c in available:
            if c.strip().lower() == alt.strip().lower(): return c
    # contains
    for alt in [target] + ALIASES.get(target, []):
        for c in available:
            if alt.lower() in c.lower(): return c
    return None

def fmts(book):
    return {
        "bold":   book.add_format({'bold': True}),
        "header": book.add_format({'bold': True,'bg_color':'#D9D9D9','border':1,'align':'center','valign':'vcenter'}),
        "cell":   book.add_format({'border': 1}),
        "date":   book.add_format({'border': 1,'num_format':'yyyy-mm-dd'}),
        "money":  book.add_format({'border': 1,'num_format':'#,##0'}),
        "center": book.add_format({'border': 1,'align':'center'}),
    }

def safe_write(ws, r, c, v, F, colname=None):
    if pd.isna(v):
        ws.write(r, c, "", F["cell"])
    elif isinstance(v, pd.Timestamp):
        ws.write_datetime(r, c, v.to_pydatetime(), F["date"])
    else:
        if colname in {"Валютын дүн","Дебет дүн","Кредит дүн","Transaction","ABS"}:
            try:
                n = float(str(v).replace(",", ""))
                ws.write_number(r, c, n, F["money"]); return
            except: pass
        ws.write(r, c, v, F["cell"])

# ---------------------------------------------------------------------
# RAW sheets
# ---------------------------------------------------------------------
def sheet_gl_raw(gl_raw: pd.DataFrame, wb, writer, src_sheet: str):
    name = clean_sheet_name("GL - Raw")
    ws = wb.add_worksheet(name); writer.sheets[name] = ws
    F = fmts(wb)
    ws.write("A1", f"GL raw data (source sheet: {src_sheet})", F["bold"])
    # header
    for j,c in enumerate(gl_raw.columns, start=1): ws.write(2, j, c, F["header"])
    # rows
    for i,row in enumerate(gl_raw.itertuples(index=False), start=3):
        for j,val in enumerate(row, start=1):
            colname = gl_raw.columns[j-1]
            safe_write(ws, i, j, val, F, colname)
    # widths
    for j,c in enumerate(gl_raw.columns, start=1):
        ws.set_column(j, j, min(max(10, len(str(c))+2), 40))

def sheet_tb_raw(tb_raw: pd.DataFrame, wb, writer, src_sheet: str):
    name = clean_sheet_name("TB - Raw")
    ws = wb.add_worksheet(name); writer.sheets[name] = ws
    F = fmts(wb)
    ws.write("A1", f"TB raw data (source sheet: {src_sheet})", F["bold"])
    for j,c in enumerate(tb_raw.columns, start=1): ws.write(2, j, c, F["header"])
    for i,row in enumerate(tb_raw.itertuples(index=False), start=3):
        for j,val in enumerate(row, start=1):
            colname = tb_raw.columns[j-1]
            safe_write(ws, i, j, val, F, colname)
    for j,c in enumerate(tb_raw.columns, start=1):
        ws.set_column(j, j, min(max(10, len(str(c))+2), 40))

# ---------------------------------------------------------------------
# Test 6 – LEN
# ---------------------------------------------------------------------
def sheet_test6_len(gl: pd.DataFrame, wb, writer):
    sheet_name = clean_sheet_name("Test 6")
    ws = wb.add_worksheet(sheet_name); writer.sheets[sheet_name] = ws
    F = fmts(wb)

    order = ["Данс","Дансны нэр","Огноо","Валют","Дебет дүн","Кредит дүн","Transaction","Гүйлгээний утга"]
    sel = {c: match_col(c, list(gl.columns)) for c in order}
    available = [c for c in order if sel[c]]
    df = gl[[sel[c] for c in available]].copy()
    df.columns = available
    if "Огноо" in df.columns:
        df["Огноо"] = pd.to_datetime(df["Огноо"], errors="coerce")
    df["LEN"] = df.get("Гүйлгээний утга","").astype(str).str.len()

    ws.write("B1","Steppe Road of Development LLC",F["bold"])
    ws.write("B2","Test 6",F["bold"])
    ws.write("B3",TITLE_DATE,F["bold"])
    ws.write("B5","Procedure",F["bold"]); ws.write("C6","Extract journal entries and compute LEN of description.")
    ws.write("B8","Summary",F["bold"])
    ws.write("B9","Testing by Audit Team?",F["header"]); ws.write("C9","No. of JE selected for testing",F["header"])
    ws.write("B10","No",F["cell"]); ws.write("C10","n.a",F["cell"])
    ws.write("B12","Comment:",F["bold"]); ws.write("C13","We calculated LEN for description.")

    start = 17
    for j,c in enumerate(list(df.columns)): ws.write(start, j+1, c, F["header"])
    for i,row in enumerate(df.itertuples(index=False), start=start+1):
        for j,(col,val) in enumerate(zip(df.columns,row), start=1):
            safe_write(ws,i,j,val,F,col)
    widths={"Данс":16,"Дансны нэр":30,"Огноо":12,"Валют":6,"Дебет дүн":16,"Кредит дүн":16,"Transaction":12,"Гүйлгээний утга":32,"LEN":6}
    for idx,name in enumerate(df.columns): ws.set_column(idx+1, idx+1, widths.get(name,14))

# ---------------------------------------------------------------------
# Non-Business Day
# ---------------------------------------------------------------------
def sheet_non_business_day(gl: pd.DataFrame, wb, writer):
    sheet_name = clean_sheet_name("Non-Business Day")
    ws = wb.add_worksheet(sheet_name); writer.sheets[sheet_name] = ws
    F=fmts(wb)

    target_cols = [
        "Данс","Дансны нэр","Огноо","Гүйлгээний дугаар","Харьцсан дансны нэр","Харьцсан данс","Day of the week",
        "Баримтын дугаар","Валют","Ханш","Валютын дүн","Дебет дүн","Кредит дүн","Transaction","Гүйлгээний утга",
        "ABS","Үүсгэсэн огноо","Day of the week /posted date/","Бүртгэсэн хэрэглэгч","Цонхны нэр"
    ]
    sel = {t: match_col(t, list(gl.columns)) for t in target_cols}
    df = pd.DataFrame({t: gl[sel[t]] if sel[t] else np.nan for t in target_cols})
    if match_col("Огноо", list(df.columns)): df["Огноо"]=pd.to_datetime(df["Огноо"], errors="coerce")

    ws.write("B1","Steppe Road of Development LLC",F["bold"])
    ws.write("B2","Non business day journals",F["bold"])
    ws.write("B3",TITLE_DATE,F["bold"])
    ws.write("B5","Procedure",F["bold"])
    ws.write("C6","Extract journal entries occurring on non-business days.")
    ws.write("B8","Summary",F["bold"]); ws.write("B9","Testing by Audit Team?",F["header"]); ws.write("C9","Number of Journal Entries extracted in each test",F["header"])
    ws.write("B10","No",F["cell"]); ws.write("C10","0",F["cell"])
    ws.write("B12","Comment:",F["bold"]); ws.write("C13","Company has weekend postings for month-end; criterion not tested.")

    start=16
    for j,c in enumerate(df.columns): ws.write(start, j, c, F["header"])
    for i,row in enumerate(df.itertuples(index=False), start=start+1):
        for j,(col,val) in enumerate(zip(df.columns,row), start=0):
            safe_write(ws,i,j,val,F,col)
    for idx,name in enumerate(df.columns): ws.set_column(idx, idx, max(12, len(name)+2))

# ---------------------------------------------------------------------
# Test 8 – Keywords
# ---------------------------------------------------------------------
def sheet_test8(gl: pd.DataFrame, wb, writer):
    sheet_name = clean_sheet_name("Test 8")
    ws = wb.add_worksheet(sheet_name); writer.sheets[sheet_name] = ws
    F = fmts(wb)
    desc_col = match_col("Гүйлгээний утга", list(gl.columns))
    if not desc_col: raise ValueError("Гүйлгээний утга багана олдсонгүй")

    keywords = [
        ("Terminate", ["Дуусгах","зогсоох","цуцлах"]),
        ("Adjust",    ["Тохируулга"]),
        ("Error",     ["Алдаа"]),
        ("Wrong",     ["Буруу"]),
        ("Revise",    ["Засах","дахин хянах","өөрчлөх"]),
        ("Буцаалт",   ["Буцаалт"]),
        ("Delete",    ["Устгах","арилгах"]),
    ]

    summary, detail = [], []
    for eng, mn in keywords:
        terms = [eng.lower()] + [m.lower() for m in mn]
        mask = gl[desc_col].astype(str).str.lower().apply(lambda s: any(t in s for t in terms))
        m = gl[mask].copy()
        summary.append({"Keyword":eng,"Keyword /Mongolia/":", ".join(mn),"Number of entries Account Name":"-",
                        "Number of entries Account Name ": len(m)})
        for _, r in m.iterrows():
            detail.append({"Данс":r.get(match_col("Данс",gl.columns),""),"Гүйлгээний утга":r.get(desc_col,""),"Per audit":eng})

    sdf = pd.DataFrame(summary)
    total = {"Keyword":"Total","Keyword /Mongolia/":"","Number of entries Account Name":"-",
             "Number of entries Account Name ": int(sdf["Number of entries Account Name "].sum())}
    sdf = pd.concat([sdf, pd.DataFrame([total])], ignore_index=True)
    ddf = pd.DataFrame(detail)

    ws.write("B1","Steppe Road of Development LLC",F["bold"]); ws.write("B2","Test 8",F["bold"]); ws.write("B3",TITLE_DATE,F["bold"])
    ws.write("B5","Procedure",F["bold"]); ws.write("C6","Extract journal entries containing keywords of interest.")
    ws.write("B8","Summary",F["bold"]); ws.write("B9","Testing by Audit Team?",F["header"]); ws.write("C9","No. of JE selected for testing",F["header"])
    ws.write("B10","Yes",F["cell"]); ws.write_number("C10", int(total["Number of entries Account Name "]), F["cell"])
    ws.write("B12","Comment",F["bold"]); ws.write("C13","Searched Onch guidance keywords.")

    srow=17
    for j,c in enumerate(sdf.columns): ws.write(srow, j+1, c, F["header"])
    for i,row in enumerate(sdf.itertuples(index=False), start=srow+1):
        for j,val in enumerate(row): ws.write(i, j+1, val, F["cell"])
    ws.set_column(1,1,18); ws.set_column(2,2,35); ws.set_column(3,4,16)

    dstart = srow + len(sdf) + 3
    ws.write(dstart,1,"All entries:",F["bold"])
    for j,c in enumerate(ddf.columns): ws.write(dstart+1, j+1, c, F["header"])
    for i,row in enumerate(ddf.itertuples(index=False), start=dstart+2):
        for j,val in enumerate(row): ws.write(i, j+1, val, F["cell"])
    ws.set_column(1,1,16); ws.set_column(2,2,50); ws.set_column(3,3,16)

# ---------------------------------------------------------------------
# Test 9 – recurring 9s (>=6)
# ---------------------------------------------------------------------

def sheet_test9(gl: pd.DataFrame, wb, writer):
    sheet_name = clean_sheet_name("Test 9")
    ws = wb.add_worksheet(sheet_name); writer.sheets[sheet_name] = ws

    # formats
    bold   = wb.add_format({'bold': True})
    header = wb.add_format({'bold': True,'bg_color': '#D9D9D9','border': 1,'align': 'center','valign': 'vcenter'})
    cell   = wb.add_format({'border': 1})
    datef  = wb.add_format({'border': 1,'num_format': 'yyyy-mm-dd'})
    money  = wb.add_format({'border': 1,'num_format': '#,##0'})

    # alias mapping
    ALIASES_LOCAL = {
        "Данс": ["Данс","Account","GL Account","Account No","Account Number"],
        "Дансны нэр": ["Дансны нэр","Account Name","GL Account Name","Name"],
        "Огноо": ["Огноо","Posted Date","Posting Date","Date","GL Date","Үүсгэсэн огноо"],
        "Гүйлгээний дугаар": ["Гүйлгээний дугаар","Document No","Voucher No","Entry No","Trans No"],
        "Харьцсан дансны нэр": ["Харьцсан дансны нэр","Counter Account Name","Offset Account Name"],
        "Харьцсан данс": ["Харьцсан данс","Counter Account","Offset Account"],
        "Валют": ["Валют","Currency","Currency Code"],
        "Ханш": ["Ханш","Exchange Rate","Rate"],
        "Валютын дүн": ["Валютын дүн","Foreign Amount","FCY Amount","Amount (FCY)"],
        "Дебет дүн": ["Дебет дүн","Debit","Debit Amount","Debit (MNT)"],
        "Кредит дүн": ["Кредит дүн","Credit","Credit Amount","Credit (MNT)"],
        "Transaction": ["Transaction","Amount","Transaction Amount","Гүйлгээний дүн"],
        "Гүйлгээний утга": ["Гүйлгээний утга","Description","Memo","Narration"],
        "ABS": ["ABS","Absolute"],
        "Бүртгэсэн хэрэглэгч": ["Бүртгэсэн хэрэглэгч","User","Posted By","Created By"],
    }
    def mcol(target, available):
        for c in available:
            if c.strip().lower() == target.strip().lower(): return c
        for alt in ALIASES_LOCAL.get(target, []):
            for c in available:
                if c.strip().lower() == alt.strip().lower(): return c
        for alt in [target] + ALIASES_LOCAL.get(target, []):
            for c in available:
                if alt.lower() in c.lower(): return c
        return None

    cols = list(gl.columns)
    sel = {}
    for t in ALIASES_LOCAL:
        m = mcol(t, cols)
        if m: sel[t] = m
    assert "Transaction" in sel, "Transaction багана олдсонгүй. ALIASES жагсаалтад өөр нэр нэмнэ үү."

    import re as _re
    def has_six_or_more_9(val):
        s = _re.sub(r"[^\d]", "", str(val))
        return _re.search(r"9{6,}", s) is not None

    mask = gl[sel["Transaction"]].apply(has_six_or_more_9)
    df_out = gl[mask].copy()

    order = ["Данс","Дансны нэр","Огноо","Гүйлгээний дугаар",
             "Харьцсан дансны нэр","Харьцсан данс","Валют","Ханш","Валютын дүн",
             "Дебет дүн","Кредит дүн","Transaction","Гүйлгээний утга","ABS","Бүртгэсэн хэрэглэгч"]
    avail = [c for c in order if c in sel]
    df_out = df_out[[sel[c] for c in avail]].copy()
    df_out.columns = avail
    if "Огноо" in df_out.columns:
        df_out["Огноо"] = pd.to_datetime(df_out["Огноо"], errors="coerce")

    ws.write("B1","Steppe Road of Development LLC",bold)
    ws.write("B2","Journal entries containing recurring digits",bold)
    ws.write("B3",TITLE_DATE,bold)
    ws.write("B5","Procedure",bold)
    ws.write("C6","Extract journal entries with more than a certain number of recurring digits")
    ws.write("B8","Summary",bold)
    ws.write("B9","Testing by Audit Team?",header)
    ws.write("C9","Number of Journal Entries extracted in each test",header)
    ws.write("B10","No",cell)
    ws.write_number("C10", int(len(df_out)), cell)
    ws.write("B12","Comment",bold)
    ws.write("C13","We searched for entries with 6 or more consecutive '9' digits in Transaction (e.g., 232,999,999).")

    start = 17
    for j, name in enumerate(df_out.columns):
        ws.write(start, j+1, name, header)

    def write_safe(r, c, v, col):
        if pd.isna(v):
            ws.write(r, c, "", cell)
        elif isinstance(v, pd.Timestamp):
            ws.write_datetime(r, c, v.to_pydatetime(), datef)
        else:
            if col in {"Валютын дүн","Дебет дүн","Кредит дүн","Transaction","ABS"}:
                try:
                    n = float(str(v).replace(",", ""))
                    ws.write_number(r, c, n, money); return
                except Exception:
                    pass
            ws.write(r, c, v, cell)

    for i, row in enumerate(df_out.itertuples(index=False), start=start+1):
        for j, (col_name, v) in enumerate(zip(df_out.columns, row), start=1):
            write_safe(i, j, v, col_name)

    widths = {"Данс":16,"Дансны нэр":28,"Огноо":12,"Гүйлгээний дугаар":14,
              "Харьцсан дансны нэр":26,"Харьцсан данс":16,"Валют":6,"Ханш":10,
              "Валютын дүн":16,"Дебет дүн":16,"Кредит дүн":16,"Transaction":16,
              "Гүйлгээний утга":32,"ABS":12,"Бүртгэсэн хэрэглэгч":18}
    for idx, name in enumerate(df_out.columns):
        ws.set_column(idx+1, idx+1, widths.get(name, 14))
def sheet_test10(gl: pd.DataFrame, wb, writer):
    sheet_name = clean_sheet_name("Test 10")
    ws = wb.add_worksheet(sheet_name); writer.sheets[sheet_name] = ws
    F=fmts(wb)
    sel={k: match_col(k, list(gl.columns)) for k in ALIASES}
    tr=sel["Transaction"]; assert tr

    def has_7plus0(x): return re.search(r"0{7,}", re.sub(r"[^\d]","",str(x))) is not None
    df = gl[gl[tr].apply(has_7plus0)].copy()

    order=["Данс","Дансны нэр","Огноо","Гүйлгээний дугаар","Харьцсан дансны нэр","Харьцсан данс","Баримт дугаар",
           "Валют","Ханш","Валютын дүн","Дебет дүн","Кредит дүн","Transaction","Гүйлгээний утга","ABS","Бүртгэсэн хэрэглэгч","Type"]
    av=[c for c in order if sel.get(c)]
    df=df[[sel[c] for c in av]].copy(); df.columns=av
    if "Огноо" in df.columns: df["Огноо"]=pd.to_datetime(df["Огноо"], errors="coerce")

    ws.write("B1","Steppe Road of Development LLC",F["bold"]); ws.write("B2","Test 10",F["bold"]); ws.write("B3",TITLE_DATE,F["bold"])
    ws.write("B5","Procedure",F["bold"]); ws.write("C6","Extract entries with 7+ consecutive zeros in Transaction.")
    ws.write("B8","Summary",F["bold"]); ws.write("B9","Testing by Audit Team?",F["header"]); ws.write("C9","No. of JE selected for testing",F["header"])
    ws.write("B10","No",F["cell"]); ws.write_number("C10", int(len(df)), F["cell"])
    ws.write("B12","Comment",F["bold"]); ws.write("C13","Rounding/pretty numbers (e.g., 10,000,000).")

    start=31; ws.write(start-2,1,"Transaction list:",F["bold"])
    for j,c in enumerate(df.columns): ws.write(start, j+1, c, F["header"])
    for i,row in enumerate(df.itertuples(index=False), start=start+1):
        for j,(col,val) in enumerate(zip(df.columns,row),start=1): safe_write(ws,i,j,val,F,col)
    widths={"Данс":16,"Дансны нэр":28,"Огноо":12,"Гүйлгээний дугаар":14,"Харьцсан дансны нэр":26,"Харьцсан данс":16,
            "Баримт дугаар":14,"Валют":6,"Ханш":10,"Валютын дүн":16,"Дебет дүн":16,"Кредит дүн":16,
            "Transaction":16,"Гүйлгээний утга":40,"ABS":12,"Бүртгэсэн хэрэглэгч":18,"Type":14}
    for idx,name in enumerate(df.columns): ws.set_column(idx+1, idx+1, widths.get(name,14))

# ---------------------------------------------------------------------
# Test 11 – Top 40 by abs(Transaction)
# ---------------------------------------------------------------------
def sheet_test11(gl: pd.DataFrame, wb, writer):
    sheet_name = clean_sheet_name("Test 11")
    ws = wb.add_worksheet(sheet_name); writer.sheets[sheet_name] = ws
    F=fmts(wb)
    sel={k: match_col(k, list(gl.columns)) for k in ALIASES}
    tr=sel["Transaction"]; assert tr
    tx = gl[tr].apply(to_number)
    df = gl.assign(__ABS__=tx.abs()).sort_values("__ABS__", ascending=False).head(40).copy()

    order=["Данс","Дансны нэр","Огноо","Гүйлгээний дугаар","Харьцсан дансны нэр","Харьцсан данс",
           "Валют","Ханш","Валютын дүн","Дебет дүн","Кредит дүн","Transaction","Гүйлгээний утга","ABS","Бүртгэсэн хэрэглэгч"]
    av=[c for c in order if sel.get(c)]
    df=df[[sel[c] for c in av]].copy(); df.columns=av
    if "Огноо" in df.columns: df["Огноо"]=pd.to_datetime(df["Огноо"], errors="coerce")

    ws.write("B1","Steppe Road of Development LLC",F["bold"]); ws.write("B2","Test 11",F["bold"]); ws.write("B3",TITLE_DATE,F["bold"])
    ws.write("B5","Procedure",F["bold"]); ws.write("C6","Extract journal entries with top X largest values by Transaction.")
    ws.write("B8","Summary",F["bold"]); ws.write("B9","Testing by Audit Team?",F["header"]); ws.write("C9","No. of JE selected for testing",F["header"])
    ws.write("B10","No",F["cell"]); ws.write_number("C10", int(len(df)), F["cell"])
    ws.write("B12","Comment",F["bold"]); ws.write("C13","Transactions above MNT threshold / top ABS values.")

    start=20; ws.write(start-2,1,"Transaction above – TOP 40",F["bold"])
    for j,c in enumerate(df.columns): ws.write(start, j+1, c, F["header"])
    for i,row in enumerate(df.itertuples(index=False), start=start+1):
        for j,(col,val) in enumerate(zip(df.columns,row),start=1): safe_write(ws,i,j,val,F,col)
    widths={"Данс":16,"Дансны нэр":28,"Огноо":12,"Гүйлгээний дугаар":14,"Харьцсан дансны нэр":26,"Харьцсан данс":16,
            "Валют":6,"Ханш":10,"Валютын дүн":16,"Дебет дүн":16,"Кредит дүн":16,
            "Transaction":16,"Гүйлгээний утга":40,"ABS":14,"Бүртгэсэн хэрэглэгч":18}
    for idx,name in enumerate(df.columns): ws.set_column(idx+1, idx+1, widths.get(name,14))

# ---------------------------------------------------------------------
# Test 15 – Revenue Top10 (5,13 debit) + график
# ---------------------------------------------------------------------
def sheet_test15_rev_top10(gl: pd.DataFrame, wb, writer, TOP_N=10):
    sheet_name = clean_sheet_name("Test 15 – Rev Top10")
    ws = wb.add_worksheet(sheet_name); writer.sheets[sheet_name] = ws
    F=fmts(wb)
    sel={k: match_col(k, list(gl.columns)) for k in ALIASES}
    acc=sel["Данс"]; debit=sel["Дебет дүн"]; assert acc and debit

    acc_digits = gl[acc].astype(str).str.replace(r"\D","",regex=True)
    is_rev = acc_digits.str.startswith(("5","13"))
    df = gl[is_rev].copy()
    df["__DEBIT__"]=df[debit].apply(to_number)
    top10 = df.sort_values("__DEBIT__", ascending=False).head(TOP_N)

    order=["Данс","Дансны нэр","Огноо","Гүйлгээний дугаар","Харьцсан дансны нэр","Харьцсан данс",
           "Валют","Ханш","Валютын дүн","Дебет дүн","Кредит дүн","Transaction","Гүйлгээний утга"]
    av=[c for c in order if sel.get(c)]
    out = top10[[sel[c] for c in av]].copy(); out.columns=av
    if "Огноо" in out.columns: out["Огноо"]=pd.to_datetime(out["Огноо"], errors="coerce")

    ws.write("B1","Steppe Road of Development LLC",F["bold"]); ws.write("B2","Test 15",F["bold"]); ws.write("B3",TITLE_DATE,F["bold"])
    ws.write("B5","Procedure",F["bold"]); ws.write("C6","Revenue (accounts starting 5 or 13) Top 10 by Debit.")
    ws.write("B8","Summary",F["bold"]); ws.write("B9","Testing by Audit Team?",F["header"]); ws.write("C9","No. of JE selected for testing",F["header"])
    ws.write("B10","No",F["cell"]); ws.write_number("C10",0,F["cell"])
    ws.write("B12","Comment",F["bold"]); ws.write("C13","Scatter uses the debit amounts.")

    # Chart data
    chart_row=19
    ws.write(chart_row,1,"Rank",F["header"]); ws.write(chart_row,2,"Дебет дүн",F["header"])
    for i,v in enumerate(top10["__DEBIT__"].tolist(), start=1):
        ws.write_number(chart_row+i,1,i,F["cell"])
        if pd.isna(v): ws.write(chart_row+i,2,"",F["cell"])
        else: ws.write_number(chart_row+i,2,float(v),F["money"])
    chart=wb.add_chart({'type':'scatter','subtype':'straight_with_markers'})
    chart.set_title({'name':'Дебет дүн'}); chart.set_y_axis({'num_format':'#,##0'})
    first,last = chart_row+1, chart_row+TOP_N
    chart.add_series({'categories':[sheet_name,first,1,last,1],
                      'values':    [sheet_name,first,2,last,2],
                      'marker':{'type':'circle'}})
    ws.insert_chart('B18', chart, {'x_scale':1.15,'y_scale':1.0})

    # Table
    t0=chart_row+TOP_N+4; ws.write(t0,1,"Debit entries to PnL in December 2024:",F["bold"])
    h=t0+2
    for j,c in enumerate(out.columns): ws.write(h, j+1, c, F["header"])
    for i,row in enumerate(out.itertuples(index=False), start=h+1):
        for j,(col,val) in enumerate(zip(out.columns,row),start=1): safe_write(ws,i,j,val,F,col)
    widths={"Данс":16,"Дансны нэр":28,"Огноо":12,"Гүйлгээний дугаар":14,"Харьцсан дансны нэр":26,"Харьцсан данс":16,"Валют":6,"Ханш":10,"Валютын дүн":16,"Дебет дүн":16,"Кредит дүн":16,"Transaction":16,"Гүйлгээний утга":40}
    for idx,name in enumerate(out.columns): ws.set_column(idx+1, idx+1, widths.get(name,14))

# ---------------------------------------------------------------------
# Test 15 – Expenses list (6/7) + Transaction chart
# ---------------------------------------------------------------------

def sheet_test15_exp_list(gl: pd.DataFrame, wb, writer, max_points=200):
    sheet_name = clean_sheet_name("Test 15 – Exp 6-7 List")
    ws = wb.add_worksheet(sheet_name); writer.sheets[sheet_name] = ws
    F=fmts(wb)

    sel = {t: match_col(t, list(gl.columns)) for t in ALIASES}
    acc = sel.get("Данс"); tr = sel.get("Transaction")
    assert acc and tr, "‘Данс’ болон ‘Transaction’ багана шаардлагатай."

    digits = gl[acc].astype(str).str.replace(r"\D","",regex=True)
    exp_df = gl[digits.str.startswith(("6","7"))].copy()

    order = ["Данс","Дансны нэр","Огноо","Гүйлгээний дугаар","Харьцсан дансны нэр","Харьцсан данс",
             "Валют","Ханш","Валютын дүн","Дебет дүн","Кредит дүн","Transaction","Гүйлгээний утга"]
    available = [c for c in order if sel.get(c)]

    exp_df["__TXN__"] = exp_df[tr].apply(to_number)
    out = exp_df.sort_values("__TXN__", ascending=False)[[sel[c] for c in available]].copy()
    out.columns = available
    if "Огноо" in out.columns:
        out["Огноо"] = pd.to_datetime(out["Огноо"], errors="coerce")

    ws.write("B1","Steppe Road of Development LLC",F["bold"]) if False else None  # keep style consistent outside

    ws.write("B1","Steppe Road of Development LLC",F["bold"]) ; ws.write("B2","Test 15",F["bold"]) ; ws.write("B3",TITLE_DATE,F["bold"]) 
    ws.write("B5","Procedure",F["bold"]) ; ws.write("C6","Expense accounts (start with 6 or 7). Scatter uses Transaction from the table below.")
    ws.write("B8","Summary",F["bold"]) ; ws.write("B9","Testing by Audit Team?",F["header"]) ; ws.write("C9","No. of JE selected for testing",F["header"]) 
    ws.write("B10","No",F["cell"]) ; ws.write_number("C10", 0, F["cell"]) 
    ws.write("B12","Comment",F["bold"]) ; ws.write("C13","Scatter uses the Transaction column from the detailed table (sorted).") 

    title_row = 18; ws.write(title_row,1,"Debit entries to PnL in December 2024:",F["bold"]) 
    head_row = title_row + 2
    for j, name in enumerate(out.columns): ws.write(head_row, j+1, name, F["header"])

    for i, row in enumerate(out.itertuples(index=False), start=head_row+1):
        for j, (col_name, val) in enumerate(zip(out.columns, row), start=1):
            safe_write(ws, i, j, val, F, col_name)

    n_rows = min(max_points, len(out))
    if n_rows > 0 and "Transaction" in out.columns:
        txn_col_idx  = list(out.columns).index("Transaction") + 1
        rank_col_idx = txn_col_idx + 1
        first_row = head_row + 1
        last_row  = head_row + n_rows

        ws.write(first_row - 1, rank_col_idx, "Rank (hidden)", F["header"]) 
        for i in range(n_rows): ws.write_number(first_row + i, rank_col_idx, i + 1, F["cell"]) 
        ws.set_column(rank_col_idx, rank_col_idx, None, None, {'hidden': True})

        chart = wb.add_chart({'type':'scatter','subtype':'straight_with_markers'})
        chart.set_title({'name':'Transaction (Rank)'}) ; chart.set_y_axis({'num_format':'#,##0'})
        chart.add_series({
            'categories': [sheet_name, first_row, rank_col_idx, last_row, rank_col_idx],
            'values':     [sheet_name, first_row, txn_col_idx,  last_row, txn_col_idx],
            'marker': {'type': 'circle'},
        })
        ws.insert_chart('B15', chart, {'x_scale':1.15,'y_scale':1.0})

    widths = {"Данс":16,"Дансны нэр":28,"Огноо":12,"Гүйлгээний дугаар":14,"Харьцсан дансны нэр":26,"Харьцсан данс":16,
              "Валют":6,"Ханш":10,"Валютын дүн":16,"Дебет дүн":16,"Кредит дүн":16,"Transaction":16,"Гүйлгээний утга":40}
    for idx, name in enumerate(out.columns): ws.set_column(idx+1, idx+1, widths.get(name, 14))
def sheet_test16_revexp(gl: pd.DataFrame, wb, writer, TOP_N=10):
    sheet_name = clean_sheet_name("Test 16 – RevExp Top10")
    ws = wb.add_worksheet(sheet_name); writer.sheets[sheet_name] = ws
    F=fmts(wb)
    sel={k: match_col(k, list(gl.columns)) for k in ALIASES}
    acc=sel["Данс"]; debit=sel["Дебет дүн"]; credit=sel["Кредит дүн"]
    assert acc and debit and credit

    acc_digits = gl[acc].astype(str).str.replace(r"\D","",regex=True)
    rev_df = gl[acc_digits.str.startswith(("5","13"))].copy()
    exp_df = gl[acc_digits.str.startswith(("6","7","8"))].copy()
    rev_df["__DEBIT__"] = rev_df[debit].apply(to_number)
    exp_df["__CREDIT__"]= exp_df[credit].apply(to_number)
    top_rev = rev_df.sort_values("__DEBIT__", ascending=False).head(TOP_N)
    top_exp = exp_df.sort_values("__CREDIT__", ascending=False).head(TOP_N)

    order=["Данс","Дансны нэр","Огноо","Гүйлгээний дугаар","Харьцсан дансны нэр","Харьцсан данс",
           "Валют","Ханш","Валютын дүн","Дебет дүн","Кредит дүн","Transaction","Гүйлгээний утга"]
    av=[c for c in order if sel.get(c)]
    rev_out=top_rev[[sel[c] for c in av]].copy(); rev_out.columns=av
    exp_out=top_exp[[sel[c] for c in av]].copy(); exp_out.columns=av
    if "Огноо" in rev_out.columns: rev_out["Огноо"]=pd.to_datetime(rev_out["Огноо"], errors="coerce")
    if "Огноо" in exp_out.columns: exp_out["Огноо"]=pd.to_datetime(exp_out["Огноо"], errors="coerce")

    ws.write("B1","Steppe Road of Development LLC",F["bold"]); ws.write("B2","Test 15",F["bold"]); ws.write("B3",TITLE_DATE,F["bold"])
    ws.write("B5","Procedure",F["bold"]); ws.write("C6","Revenue (5,13) Top 10 by Debit and Expense (6,7,8) Top 10 by Credit.")
    ws.write("B8","Summary",F["bold"]); ws.write("B9","Testing by Audit Team?",F["header"]); ws.write("C9","No. of JE selected for testing",F["header"])
    ws.write("B10","No",F["cell"]); ws.write_number("C10",0,F["cell"])
    ws.write("B12","Comment",F["bold"]); ws.write("C13","Two scatter plots: debit(Rev) and credit(Exp).")

    # Revenue chart
    row_rev=19
    ws.write(row_rev,1,"Rank",F["header"]); ws.write(row_rev,2,"Дебет дүн",F["header"])
    for i,v in enumerate(top_rev["__DEBIT__"].tolist(), start=1):
        ws.write_number(row_rev+i,1,i,F["cell"])
        if pd.isna(v): ws.write(row_rev+i,2,"",F["cell"])
        else: ws.write_number(row_rev+i,2,float(v),F["money"])
    chart_rev=wb.add_chart({'type':'scatter','subtype':'straight_with_markers'})
    chart_rev.set_title({'name':'Орлого (5,13) — Дебет Top 10'}); chart_rev.set_y_axis({'num_format':'#,##0'})
    first,last=row_rev+1,row_rev+TOP_N
    chart_rev.add_series({'categories':[sheet_name,first,1,last,1],
                          'values':    [sheet_name,first,2,last,2],
                          'marker':{'type':'circle'}})
    ws.insert_chart('B18', chart_rev, {'x_scale':1.15,'y_scale':1.0})

    # Expense chart
    row_exp = row_rev + TOP_N + 16
    ws.write(row_exp,1,"Rank",F["header"]); ws.write(row_exp,2,"Кредит дүн",F["header"])
    for i,v in enumerate(top_exp["__CREDIT__"].tolist(), start=1):
        ws.write_number(row_exp+i,1,i,F["cell"])
        if pd.isna(v): ws.write(row_exp+i,2,"",F["cell"])
        else: ws.write_number(row_exp+i,2,float(v),F["money"])
    chart_exp=wb.add_chart({'type':'scatter','subtype':'straight_with_markers'})
    chart_exp.set_title({'name':'Зардал (6,7,8) — Кредит Top 10'}); chart_exp.set_y_axis({'num_format':'#,##0'})
    first2,last2=row_exp+1,row_exp+TOP_N
    chart_exp.add_series({'categories':[sheet_name,first2,1,last2,1],
                          'values':    [sheet_name,first2,2,last2,2],
                          'marker':{'type':'circle'}})
    ws.insert_chart(f'B{row_exp-1}', chart_exp, {'x_scale':1.15,'y_scale':1.0})

    # Tables (rev, exp)
    t1=row_exp+TOP_N+6; ws.write(t1,1,"Revenue (5,13) — Debit Top 10:",F["bold"])
    h1=t1+2
    for j,c in enumerate(rev_out.columns): ws.write(h1, j+1, c, F["header"])
    for i,row in enumerate(rev_out.itertuples(index=False), start=h1+1):
        for j,(col,val) in enumerate(zip(rev_out.columns,row),start=1): safe_write(ws,i,j,val,F,col)

    t2=h1+TOP_N+5; ws.write(t2,1,"Expense (6,7,8) — Credit Top 10:",F["bold"])
    h2=t2+2
    for j,c in enumerate(exp_out.columns): ws.write(h2, j+1, c, F["header"])
    for i,row in enumerate(exp_out.itertuples(index=False), start=h2+1):
        for j,(col,val) in enumerate(zip(exp_out.columns,row),start=1): safe_write(ws,i,j,val,F,col)

    widths={"Данс":16,"Дансны нэр":28,"Огноо":12,"Гүйлгээний дугаар":14,"Харьцсан дансны нэр":26,"Харьцсан данс":16,"Валют":6,"Ханш":10,"Валютын дүн":16,"Дебет дүн":16,"Кредит дүн":16,"Transaction":16,"Гүйлгээний утга":40}
    for nm,w in widths.items():
        if nm in rev_out.columns: ws.set_column(list(rev_out.columns).index(nm)+1, list(rev_out.columns).index(nm)+1, w)
        if nm in exp_out.columns: ws.set_column(list(exp_out.columns).index(nm)+1, list(exp_out.columns).index(nm)+1, w)

# ---------------------------------------------------------------------
# Main – add RAW sheets at the end
# ---------------------------------------------------------------------
def main():
    # Load GL & TB
    gl_raw, gl_src = load_gl(INPUT_XLSX_GL)
    try:
        tb_raw, tb_src = load_tb(INPUT_XLSX_TB)
    except Exception:
        tb_raw, tb_src = pd.DataFrame(), "(not found)"

    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        wb = writer.book

        # ======== 1) 9 TEST SHEETS ========
        sheet_test6_len(gl_raw, wb, writer)
        sheet_non_business_day(gl_raw, wb, writer)
        sheet_test8(gl_raw, wb, writer)
        sheet_test9(gl_raw, wb, writer)
        sheet_test10(gl_raw, wb, writer)
        sheet_test11(gl_raw, wb, writer)
        sheet_test15_rev_top10(gl_raw, wb, writer)
        sheet_test15_exp_list(gl_raw, wb, writer)
        sheet_test16_revexp(gl_raw, wb, writer)

        # ======== 2) RAW SHEETS (always at the back) ========
        sheet_gl_raw(gl_raw, wb, writer, gl_src)
        if not tb_raw.empty:
            sheet_tb_raw(tb_raw, wb, writer, tb_src)

    print(f"✔ Done. Workbook written to: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
