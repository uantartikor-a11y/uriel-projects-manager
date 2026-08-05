from datetime import datetime
import io
import os
import pandas as pd
import streamlit as st

SYSTEM_USER = "admin"
SYSTEM_PASSWORD = "Plm753&%#"

st.set_page_config(
    page_title="מערכת ניהול פרויקטים ודוחות", page_icon="📊", layout="wide"
)

def check_login():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if st.session_state["logged_in"]:
        return True

    st.title("🔒 כניסה למערכת ניהול פרויקטים")
    st.write("נא להזין שם משתמש וסיסמה כדי לגשת למערכת הנתונים.")
    
    with st.form("login_form"):
        entered_user = st.text_input("שם משתמש:")
        entered_password = st.text_input("סיסמה:", type="password")
        submit_btn = st.form_submit_button("התחבר")
        
        if submit_btn:
            if entered_user == SYSTEM_USER and entered_password == SYSTEM_PASSWORD:
                st.session_state["logged_in"] = True
                st.rerun()
            else:
                st.error("שם משתמש או סיסמה שגוים, נסה שוב.")
                
    return False

if not check_login():
    st.stop()

st.markdown(
    """
    <style>
    .main, .stSidebar, body, p, h1, h2, h3, h4, h5, h6, span, div {
        direction: rtl !important;
        text-align: right !important;
    }
    .stApp {
        background-color: #f8fafc;
    }
    h1 {
        color: #0f766e !important;
        font-weight: 800 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 מערכת ניהול תקציב, פרויקטים ודוחות")

EXCEL_FILE = "DATE.xlsx"

try:
    if os.path.exists(EXCEL_FILE):
        xls = pd.ExcelFile(EXCEL_FILE)
        st.success(f"קובץ הנתונים נטען בהצלחה! גיליונות זמינים: {', '.join(xls.sheet_names)}")
        
        sheet_choice = st.selectbox("בחר גיליון לצפייה:", xls.sheet_names)
        df = pd.read_excel(xls, sheet_name=sheet_choice)
        st.dataframe(df, use_container_width=True)
    else:
        st.error(f"הקובץ {EXCEL_FILE} לא נמצא ברפוזיטורי.")
except Exception as e:
    st.error(f"שגיאה בטעינת הנתונים: {e}")

if st.sidebar.button("🚪 התנתק"):
    st.session_state["logged_in"] = False
    st.rerun()
