import streamlit as st
import pandas as pd
import os

SYSTEM_USER = "admin"
SYSTEM_PASSWORD = "Plm753&%#"

st.set_page_config(page_title="מערכת ניהול פרויקטים ודוחות", page_icon="📊", layout="wide")

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

st.markdown("""
    <style>
    .main, .stSidebar, body, p, h1, h2, h3, h4, h5, h6, span, div {
        direction: rtl !important;
        text-align: right !important;
    }
    .stApp { background-color: #f8fafc; }
    h1, h2, h3 { color: #0f766e !important; font-weight: 800 !important; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 מערכת ניהול תקציב, פרויקטים ודוחות")

EXCEL_FILE = "DATE.xlsx"

if os.path.exists(EXCEL_FILE):
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        st.success("✅ קובץ הנתונים מהענן מחובר בהצלחה!")
        
        sheet_choice = st.selectbox("בחר גיליון לצפייה:", xls.sheet_names)
        
        # קריאת הגיליון והסרת עמודות וורות לחלוטין
        df = pd.read_excel(xls, sheet_name=sheet_choice)
        df = df.dropna(how='all', axis=1) # מחיקת עמודות ריקות לחלוטין
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')] # מחיקת עמודות ללא כותרת
        
        st.dataframe(df, use_container_width=True)
        
    except Exception as e:
        st.error(f"שגיאה בניתוח קובץ האקסל: {e}")
else:
    st.error(f"הקובץ {EXCEL_FILE} אינו נמצא ברפוזיטורי.")

if st.sidebar.button("🚪 התנתק"):
    st.session_state["logged_in"] = False
    st.rerun()
