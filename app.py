# app.py
import streamlit as st
import hashlib
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv

# .env файлаас environment variable-уудыг унших
load_dotenv()

st.set_page_config(page_title="JET App Suite", page_icon="📊", layout="wide")

# =========================
# 🔐 Config
# =========================
# Нууц үгийн MD5 хэшийг environment variable-аас унших
HASH_HEX = os.getenv("JET_APP_HASH", "").strip().lower()

if not HASH_HEX:
    st.error("❌ JET_APP_HASH environment variable тохируулаагүй байна. Та .env файл үүсгэж, JET_APP_HASH-ийг тохируулна уу.")
    st.stop()

# Брютфорсын хамгаалалт
MAX_FAILS = 5
LOCK_MINUTES = 3

def _is_locked():
    fails = st.session_state.get("fail_count", 0)
    until = st.session_state.get("lock_until")
    if until and datetime.utcnow() < until:
        return True, (until - datetime.utcnow()).seconds
    if fails >= MAX_FAILS:
        st.session_state["lock_until"] = datetime.utcnow() + timedelta(minutes=LOCK_MINUTES)
        st.session_state["fail_count"] = 0
        return True, LOCK_MINUTES * 60
    return False, 0

def _reset_lock():
    if "fail_count" in st.session_state:
        del st.session_state["fail_count"]
    if "lock_until" in st.session_state:
        del st.session_state["lock_until"]

def check_password() -> bool:
    # ✅ Force logout via URL: ?logout=1
    params = st.query_params
    if params.get("logout") == "1":
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

    # Already signed in?
    if st.session_state.get("auth_ok"):
        with st.sidebar:
            st.caption("✅ Нэвтэрсэн")
            if st.button("🚪 Системээс гарах"):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
        return True

    # Lock state?
    locked, left = _is_locked()
    
    # Нэвтрэлтийн хэсгийг төвд харуулах
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.title("JET App Suite")
        st.write("Энд нэвтэрч байж цааш үргэлжилнэ.")

        if locked:
            minutes = left // 60
            seconds = left % 60
            st.error(f"⏳ Түр түгжээтэй байна. Дахин оролдох боломж {minutes} мин {seconds} сек дараа.")
            st.stop()

        with st.form("login", clear_on_submit=True):
            pwd = st.text_input("Нууц үг", type="password")
            ok = st.form_submit_button("Нэвтрэх")

        if ok:
            # Нууц үг шалгах
            entered_hash = hashlib.md5((pwd or "").encode()).hexdigest().lower()
            if entered_hash == HASH_HEX:
                st.session_state["auth_ok"] = True
                _reset_lock()
                st.rerun()
            else:
                current_fails = st.session_state.get("fail_count", 0) + 1
                st.session_state["fail_count"] = current_fails
                remaining = MAX_FAILS - current_fails
                st.error(f"❌ Буруу нууц үг байна. Дараах оролдлого: {remaining}")
    
    st.stop()

# =========================
# 🧭 Sidebar Nav
# =========================
def sidebar_nav() -> str:
    with st.sidebar:
        st.markdown("### Pages")
        return st.radio(
            "Хуудас сонгох",
            ["Нүүр", "JET Additional", "JET Statistical"],
            index=0,
            label_visibility="collapsed",
        )

# =========================
# Pages
# =========================
def page_home():
    st.title("JET App Suite")
    st.write("Зүүн талын Pages хэсгээс сонгоод аппаа ажиллуулна уу.")
    with st.expander("Тайлбар / Usage", expanded=True):
        st.markdown(
            """
            - Энэ хувилбар **нэг файл** (`app.py`) доторх олон хуудсыг sidebar-аас сонгож харуулдаг.
            - Нэвтрэлт нь **бүх хуудсанд** жигд үйлчилнэ.
            - Гарах бол sidebar дахь **"🚪 Системээс гарах"** товч эсвэл URL дээр `?logout=1`.
            - Анхны нууц үг: `Jet test16`
            """
        )

def page_jet_additional():
    st.header("JET Additional")
    st.info("Энд таны 'JET Additional' аппын логик/контентийг байрлуулна.")

def page_jet_statistical():
    st.header("JET Statistical")
    st.info("Энд таны 'JET Statistical' аппын логик/контентийг байрлуулна.")

# =========================
# 🚀 Entry
# =========================
def main():
    if not check_password():
        return
    page = sidebar_nav()
    if page == "Нүүр":
        page_home()
    elif page == "JET Additional":
        page_jet_additional()
    elif page == "JET Statistical":
        page_jet_statistical()
    else:
        page_home()

if __name__ == "__main__":
    main()
