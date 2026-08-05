import streamlit as st
import pandas as pd
import traceback

st.set_page_config(page_title="ניהול פרויקטים", page_icon="📊", layout="wide")

st.title("📊 מערכת ניהול תקציב, פרויקטים ודוחות")

try:
    EXCEL_FILE = "DATE.xlsx"
    xls = pd.ExcelFile(EXCEL_FILE)
    st.success(f"הקובץ נטען בהצלחה! גיליונות: {xls.sheet_names}")
    
    sheet_choice = st.selectbox("בחר גיליון לצפייה:", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet_choice)
    st.dataframe(df, use_container_width=True)
except Exception as e:
    st.error("שגיאה מפורטת בהרצת המערכת:")
    st.code(traceback.format_exc())
