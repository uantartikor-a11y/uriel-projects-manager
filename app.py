from datetime import datetime
import io
import os
import pandas as pd
import streamlit as st

SYSTEM_USER = "admin"
SYSTEM_PASSWORD = "Plm753&%#"

st.set_page_config(
    page_title="ניהול פרויקטים ודוחות", page_icon="📊", layout="wide"
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

EXCEL_FILE = "פרויקטים אוריאל אלמוג.xlsx"

def load_data():
    if not os.path.exists(EXCEL_FILE):
        st.error(f"קובץ האקסל '{EXCEL_FILE}' אינו נמצא ברפוזיטורי!")
        st.stop()
        
    xls = pd.ExcelFile(EXCEL_FILE)
    if 'סיכום קבלת כסף' in xls.sheet_names:
        df_sum = pd.read_excel(xls, sheet_name='סיכום קבלת כסף')
        
        summary_contracts = {}
        summary_received = {}
        status_mapping = {}
        
        for i, r in df_sum.iterrows():
            if i >= 2:
                p_name_1 = r.iloc[2]
                val_contract = r.iloc[3]
                stat = r.iloc[4]
                if pd.notna(p_name_1):
                    p_str = str(p_name_1).strip()
                    if pd.notna(val_contract) and isinstance(val_contract, (int, float)):
                        summary_contracts[p_str] = float(val_contract)
                    if pd.notna(stat):
                        status_mapping[p_str] = str(stat).strip()
                
                p_name_2 = r.iloc[8]
                val_received = r.iloc[7]
                if pd.notna(p_name_2) and pd.notna(val_received) and isinstance(val_received, (int, float)):
                    p_str2 = str(p_name_2).strip()
                    summary_received[p_str2] = summary_received.get(p_str2, 0.0) + float(val_received)

        projects_rows = []
        proj_idx = 1
        
        for sheet_name in xls.sheet_names:
            if sheet_name == 'סיכום קבלת כסף':
                continue
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            p_name = sheet_name
            for i, r in df.iterrows():
                for c in range(len(r)):
                    val = str(r.iloc[c])
                    if "שם הפרוייקט:" in val or "שם הפרוייקט" in val:
                        for nc in range(c+1, len(r)):
                            if pd.notna(r.iloc[nc]) and str(r.iloc[nc]).strip() != '':
                                p_name = str(r.iloc[nc]).strip()
                                break
                        break
                else:
                    continue
                break
                
            p_code = f"PRJ-{proj_idx:03d}"
            proj_idx += 1
            
            contract_amount = summary_contracts.get(p_name, summary_contracts.get(sheet_name, 0.0))
            paid_amount = summary_received.get(p_name, summary_received.get(sheet_name, 0.0))
            status = status_mapping.get(p_name, status_mapping.get(sheet_name, "בביצוע"))

            projects_rows.append({
                "קוד פרויקט": p_code,
                "שם פרויקט": p_name,
                "סכום חוזה": contract_amount,
                "שולם בפועל": paid_amount,
                "סטטוס": status,
                "תאריך פתיחה": "2026-01-01",
                "תאריך סגירה": ""
            })

        df_proj = pd.DataFrame(projects_rows)
    else:
        df_proj = pd.read_excel(xls, sheet_name=0)
        
    return df_proj

try:
    df_projects = load_data()
except Exception as e:
    st.error(f"שגיאה בטעינת הנתונים: {e}")
    st.stop()

st.success("המערכת מחוברת וטעינת הנתונים הצליחה!")
st.dataframe(df_projects, use_container_width=True)

if st.sidebar.button("🚪 התנתק"):
    st.session_state["logged_in"] = False
    st.rerun()
