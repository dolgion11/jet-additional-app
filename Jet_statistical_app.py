# app.py
import streamlit as st
from pathlib import Path
import importlib
import inspect

st.set_page_config(page_title="JET Audit Automation", layout="centered")
st.title("📊 JET Audit Automation Report Generator")

st.markdown("""
Энэхүү апп нь таны оруулсан **GL / TB Excel файлуудаас**
автомат аудиторын тестийн тайлан үүсгэнэ.
""")

# -----------------------------
# report модулийг уян хатан ачаалах
# -----------------------------
def load_report_module():
    # 1) Таны өмнө хэлсэн 2 нэрийг дарааллаар оролдоно
    candidates = ["all_reports_master_merged", "build_full_report_pretty"]
    errors = []
    for name in candidates:
        try:
            return importlib.import_module(name), name
        except Exception as e:
            errors.append(f"{name}: {e}")
    raise ImportError("report модулийг ачаалж чадсангүй:\n" + "\n".join(errors))

report, report_module_name = load_report_module()

# -----------------------------
# Модулийн callable-уудыг жагсаах
# -----------------------------
def list_callables(mod):
    items = []
    for k, v in vars(mod).items():
        if callable(v) and not k.startswith("_"):
            try:
                sig = str(inspect.signature(v))
            except Exception:
                sig = "(unknown)"
            items.append((k, sig))
    # нэрээр нь багахан эрэмбэлнэ (report/build/run гэх мэт дээр ирэхэд тус болно)
    def _score(n):
        base = 0
        if "report" in n: base -= 3
        if "build"  in n: base -= 2
        if "run"    in n: base -= 1
        if "main"   in n: base -= 4
        return (base, n)
    items.sort(key=lambda x: _score(x[0]))
    return items

available_funcs = list_callables(report)

with st.expander("⚙ Илэрсэн функцууд (модулиас)"):
    if available_funcs:
        st.write(
            "\n".join([f"- `{n}{sig}`" for n, sig in available_funcs])
        )
    else:
        st.warning("Callable функц олдсонгүй. Модулийнхоо дотор public функц нэмээрэй (ж. `main`, `build_report`, `run`, …).")

# Хэрэглэгчээс функцийн нэр авах (сонгосон эсвэл гараар бичих)
default_func = available_funcs[0][0] if available_funcs else ""
func_name = st.selectbox(
    "Дуудах функцээ сонго (эсвэл доор гараар бич):",
    [default_func] + [n for n, _ in available_funcs if n != default_func]
) if available_funcs else ""

func_name_manual = st.text_input("Гараар функцийн нэр оруулах (заавал биш):", value="")
entry_name = func_name_manual.strip() or func_name.strip()

# -----------------------------
# Файл upload
# -----------------------------
gl_file = st.file_uploader("GL Excel файл оруулна уу", type=["xlsx"])
tb_file = st.file_uploader("TB Excel файл оруулж болно (заавал биш)", type=["xlsx"])

# -----------------------------
# Туслах: entry-г уян хатан дуудах
# -----------------------------
def run_entry(entry, gl_path: Path, tb_path: Path | None, out_path: Path):
    # Зарим код глобал хувьсагчид түшиглэдэг байж болох тул урьдчилан онооно
    setattr(report, "INPUT_XLSX_GL", gl_path)
    setattr(report, "INPUT_XLSX_TB", tb_path or gl_path)
    setattr(report, "OUTPUT_XLSX", out_path)

    # Параметрийн хөрвүүлэлт
    tb_eff = tb_path or gl_path

    # signature харж уян хатан дуудна
    try:
        sig = inspect.signature(entry)
        params = list(sig.parameters.keys())
    except Exception:
        params = []

    try:
        if len(params) >= 3:
            return entry(gl_path, tb_eff, out_path)
        elif len(params) == 2:
            return entry(gl_path, tb_eff)
        elif len(params) == 1:
            return entry(gl_path)
        else:
            return entry()
    except TypeError:
        # өөр тохиолдолд ээлжлэн оролдъё
        for args in [
            (gl_path, tb_eff, out_path),
            (gl_path, tb_eff),
            (gl_path,),
            tuple()
        ]:
            try:
                return entry(*args)
            except TypeError:
                continue
        raise

# -----------------------------
# Үүсгэх товч
# -----------------------------
if st.button("✅ Тайлан үүсгэх"):
    if not gl_file:
        st.error("⚠ GL файл заавал оруулах хэрэгтэй.")
    elif not entry_name:
        st.error("⚠ Дуудах функцийн нэрээ сонгох/бичих хэрэгтэй.")
    else:
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
            entry = getattr(report, entry_name)
        except AttributeError:
            st.error(f"❌ `{report_module_name}` модульд `{entry_name}` нэртэй функц байхгүй байна.")
        else:
            try:
                with st.spinner(f"⏳ `{entry_name}` ажиллаж байна..."):
                    run_entry(entry, gl_path, tb_path, out_path)

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
