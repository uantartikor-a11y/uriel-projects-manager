import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="ניהול פרויקטים", page_icon="📊", layout="wide")

st.title("📊 מערכת ניהול תקציב, פרויקטים ודוחות")

EXCEL_FILE = "פרויקטים אוריאל אלמוג.xlsx"

if os.path.exists(EXCEL_FILE):
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        st.success(f"קובץ הנתונים נטען בהצלחה! גיליונות זמינים: {', '.join(xls.sheet_names)}")
        
        sheet_choice = st.selectbox("בחר גיליון לצפייה:", xls.sheet_names)
        df = pd.read_excel(xls, sheet_name=sheet_choice)
        st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.error(f"שגיאה בקריאת קובץ האקסל: {e}")
else:
    st.error(f"קובץ האקסל '{EXCEL_FILE}' לא נמצא ברפוזיטורי. ודא ששמו מופיע בדיוק כמו שצריך.")
