from datetime import datetime
import io
import os
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT

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

def fix_hebrew(text):
    if not isinstance(text, str):
        return str(text)
    words = text.split(' ')
    fixed_words = []
    for word in words:
        has_open = '(' in word
        has_close = ')' in word
        clean_word = word.replace('(', '').replace(')', '').replace('[', '').replace(']', '')
        if any(ord('א') <= ord(c) <= ord('ת') or '\u0590' <= c <= '\u05FF' for c in clean_word):
            reversed_word = clean_word[::-1]
        else:
            reversed_word = clean_word
        if has_open and has_close:
            final_word = f"({reversed_word})"
        elif has_open:
            final_word = f"){reversed_word}"
        elif has_close:
            final_word = f"({reversed_word}"
        else:
            final_word = reversed_word
        fixed_words.append(final_word)
    return ' '.join(fixed_words[::-1])

# רישום גופן עברית תומך ל־PDF
PDF_FONT = "Helvetica"
try:
    font_path = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts\\arial.ttf")
    if not os.path.exists(font_path):
        font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    pdfmetrics.registerFont(TTFont("HebrewArial", font_path))
    PDF_FONT = "HebrewArial"
except Exception:
    pass

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
    h2, h3 {
        color: #115e59 !important;
    }
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, #0f766e 0%, #14b8a6 100%) !important;
        padding: 20px !important;
        border-radius: 14px !important;
        box-shadow: 0 6px 12px rgba(15, 118, 110, 0.2) !important;
        border: none !important;
    }
    div[data-testid="stMetric"] label {
        color: #ccfbf1 !important;
        font-weight: 600 !important;
        font-size: 15px !important;
    }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #ffffff !important;
        font-weight: 800 !important;
        font-size: 26px !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #0f172a !important;
    }
    section[data-testid="stSidebar"] p, 
    section[data-testid="stSidebar"] span, 
    section[data-testid="stSidebar"] label, 
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3 {
        color: #f1f5f9 !important;
    }
    .stButton>button {
        background-color: #0f766e;
        color: white;
        border-radius: 8px;
        border: none;
        font-weight: 600;
        padding: 0.5rem 1rem;
    }
    .stButton>button:hover {
        background-color: #115e59;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 מערכת ניהול תקציב, פרויקטים ודוחות")
st.markdown("ניהול חוזים, תשלומים, יתרות, ספקים, מע\"מ (18%) והפקדת דוחות PDF/Excel.")

EXCEL_FILE = "projects_data.xlsx"
ORIGINAL_EXCEL = "DATE.xlsx"

def load_data():
    if not os.path.exists(EXCEL_FILE) and os.path.exists(ORIGINAL_EXCEL):
        xls = pd.ExcelFile(ORIGINAL_EXCEL)
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
        expenses_rows = []
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
            if contract_amount == 0.0:
                for i, r in df.iterrows():
                    row_str = " ".join([str(v) for v in r.values if pd.notna(v)])
                    if "מחיר הפרוייקט לפי חוזה" in row_str or 'מחיר מצטבר לא כולל מע"מ' in row_str:
                        for val in r.values:
                            if isinstance(val, (int, float)) and val > 0:
                                contract_amount = float(val)
                                break

            paid_amount = summary_received.get(p_name, summary_received.get(sheet_name, 0.0))
            status = status_mapping.get(p_name, status_mapping.get(sheet_name, "בביצוע"))

            today_date = "2026-01-01"
            close_date = datetime.now().strftime("%Y-%m-%d") if status in ["נמסר", "סגור"] else ""

            projects_rows.append({
                "קוד פרויקט": p_code,
                "שם פרויקט": p_name,
                "סכום חוזה": contract_amount,
                "שולם בפועל": paid_amount,
                "סטטוס": status,
                "תאריך פתיחה": today_date,
                "תאריך סגירה": close_date
            })

            exp_started = False
            current_category = ""
            seen_exp = set()
            for i, r in df.iterrows():
                row_vals = [str(v) for v in r.values if pd.notna(v)]
                row_str = " ".join(row_vals)
                if "הוצאות פרוייקט" in row_str:
                    exp_started = True
                    continue
                if exp_started:
                    if len(row_vals) >= 2 and row_vals[0] == "תשלום":
                        current_category = row_vals[1]
                    elif len(row_vals) >= 1 and row_vals[0] not in ["תשלום", "סה\"כ", "סיכום הוצאות", "סה\"כ רווח"]:
                        for v in r.values:
                            if isinstance(v, (int, float)) and v > 0:
                                cat_name = current_category if current_category else "הוצאה שוטפת"
                                exp_key = (p_code, cat_name, float(v))
                                if exp_key not in seen_exp:
                                    seen_exp.add(exp_key)
                                    expenses_rows.append({
                                        "קוד פרויקט": p_code,
                                        "ספק / תחום": cat_name,
                                        "סכום": float(v)
                                    })
                                break

        df_proj = pd.DataFrame(projects_rows)
        df_exp = pd.DataFrame(expenses_rows)
        
        df_exp = df_exp.drop_duplicates(subset=["קוד פרויקט", "ספק / תחום", "סכום"]).reset_index(drop=True)
        
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df_proj.to_excel(writer, sheet_name='פרויקטים', index=False)
            df_exp.to_excel(writer, sheet_name='הוצאות', index=False)

    if os.path.exists(EXCEL_FILE):
        df_expenses = pd.read_excel(EXCEL_FILE, sheet_name="הוצאות")
        df_projects = pd.read_excel(EXCEL_FILE, sheet_name="פרויקטים")
    else:
        st.error(f"קובץ הנתונים {EXCEL_FILE} לא נמצא.")
        st.stop()

    if "סטטוס" not in df_projects.columns:
        df_projects["סטטוס"] = "בביצוע"
    if "תאריך פתיחה" not in df_projects.columns:
        df_projects["תאריך פתיחה"] = datetime.now().strftime("%Y-%m-%d")
    if "תאריך סגירה" not in df_projects.columns:
        df_projects["תאריך סגירה"] = ""
    else:
        df_projects["תאריך סגירה"] = df_projects["תאריך סגירה"].fillna("").astype(str)
    if "שולם בפועל" not in df_projects.columns:
        df_projects["שולם בפועל"] = 0.0
    if "ספק / תחום" not in df_expenses.columns:
        df_expenses["ספק / תחום"] = ""
    if "סכום" not in df_expenses.columns:
        df_expenses["סכום"] = 0.0

    return df_projects, df_expenses

try:
    df_projects, df_expenses = load_data()
except Exception as e:
    st.error(f"שגיאה בטעינת הקובץ: {e}")
    st.stop()

vat_rate = 0.18

def to_excel_download(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output.getvalue()

st.sidebar.header("📁 ניהול מערכת")

if st.sidebar.button("🚪 התנתק מהמערכת"):
    st.session_state["logged_in"] = False
    st.rerun()

st.sidebar.divider()

view_mode = st.sidebar.radio(
    "בחר תצוגה:",
    ["🔍 ניהול פרויקט ספציפי", "📈 סיכום כללי (דשבורד עסק)"],
)

st.sidebar.divider()

with st.sidebar.form("new_project_form"):
    st.subheader("➕ הוספת פרויקט / מכרז חדש")
    new_p_code = st.text_input("קוד פרויקט (למשל: PRJ-018)")
    new_p_name = st.text_input("שם פרויקט")
    new_p_amount = st.number_input("סכום חוזה (ללא מע\"מ)", min_value=0.0, step=1000.0)
    new_p_paid = st.number_input("תקבולים בפועל (ללא מע\"מ)", min_value=0.0, step=1000.0)
    new_p_status = st.selectbox(
        "סטטוס התחלתי:", ["מכרז", "בביצוע", "מעוכב", "בבדיקה", "נמסר", "סגור"]
    )

    add_project_submitted = st.form_submit_button("צור פרויקט")

    if add_project_submitted and new_p_code and new_p_name:
        if new_p_code in df_projects["קוד פרויקט"].values:
            st.sidebar.error("קוד פרויקט זה כבר קיים!")
        else:
            today_date = datetime.now().strftime("%Y-%m-%d")
            initial_close_date = today_date if new_p_status in ["נמסר", "סגור"] else ""

            new_proj_row = pd.DataFrame(
                {
                    "קוד פרויקט": [new_p_code],
                    "שם פרויקט": [new_p_name],
                    "סכום חוזה": [new_p_amount],
                    "שולם בפועל": [new_p_paid],
                    "סטטוס": [new_p_status],
                    "תאריך פתיחה": [today_date],
                    "תאריך סגירה": [initial_close_date],
                }
            )
            df_projects = pd.concat([df_projects, new_proj_row], ignore_index=True)

            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                df_projects.to_excel(writer, sheet_name="פרויקטים", index=False)
                df_expenses.to_excel(writer, sheet_name="הוצאות", index=False)

            st.sidebar.success(f"הפרויקט '{new_p_name}' נוסף בהצלחה!")
            st.rerun()

if view_mode == "📈 סיכום כללי (דשבורד עסק)":
    st.header("📈 סיכום פיננסי כללי לכל הפרויקטים והמכרזים")

    if df_projects.empty:
        st.info("אין עדיין פרויקטים במערכת.")
    else:
        summary_data = []
        for index, row in df_projects.iterrows():
            p_code = row["קוד פרויקט"]
            p_name = row["שם פרויקט"]
            c_amount = row["סכום חוזה"]
            paid_amount = row["שולם בפועל"] if "שולם בפועל" in df_projects.columns else 0
            status = row["סטטוס"] if "סטטוס" in df_projects.columns else "בביצוע"
            open_date = row["תאריך פתיחה"] if "תאריך פתיחה" in df_projects.columns else ""
            close_date = row["תאריך סגירה"] if "תאריך סגירה" in df_projects.columns else ""

            c_amount_vat = c_amount * (1 + vat_rate)
            paid_amount_vat = paid_amount * (1 + vat_rate)
            remaining_vat = c_amount_vat - paid_amount_vat

            p_exp = (
                df_expenses[df_expenses["קוד פרויקט"] == p_code]["סכום"].sum()
                if not df_expenses.empty
                else 0
            )
            p_net_profit = c_amount - p_exp

            summary_data.append(
                {
                    "קוד פרויקט": p_code,
                    "שם פרויקט": p_name,
                    "סטטוס": status,
                    "תאריך פתיחה": open_date,
                    "תאריך סגירה": close_date,
                    "סכום חוזה (ללא מע\"מ)": c_amount,
                    "תקבולים בפועל (ללא מע\"מ)": paid_amount,
                    "יתרה לתשלום (כולל מע\"מ)": remaining_vat,
                    "סה\"כ הוצאות": p_exp,
                    "רווח נקי": p_net_profit,
                }
            )

        df_summary = pd.DataFrame(summary_data)

        tot_contracts = df_summary["סכום חוזה (ללא מע\"מ)"].sum()
        tot_paid = df_summary["תקבולים בפועל (ללא מע\"מ)"].sum()
        tot_remaining_vat = df_summary["יתרה לתשלום (כולל מע\"מ)"].sum()
        tot_exp_all = df_summary["סה\"כ הוצאות"].sum()
        tot_net_profit = df_summary["רווח נקי"].sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("סה\"כ חוזים (ללא מע\"מ)", f"{tot_contracts:,.0f} ₪")
        c2.metric("סה\"כ תקבולים בפועל", f"{tot_paid:,.0f} ₪")
        c3.metric("סה\"כ הוצאות", f"{tot_exp_all:,.0f} ₪")
        c4.metric("סה\"כ רווח נקי", f"{tot_net_profit:,.0f} ₪")

        st.divider()
        st.subheader("📥 ייצוא והפקדת דוחות (כלל הפרויקטים)")
        col_dl1, col_dl2 = st.columns(2)

        with col_dl1:
            excel_data = to_excel_download({"פרויקטים": df_projects, "הוצאות": df_expenses})
            st.download_button(
                label="📥 הורד קובץ Excel מלא (בעברית מלאה)",
                data=excel_data,
                file_name="All_Projects_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with col_dl2:
            def generate_general_pdf(df):
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                elements = []
                styles = getSampleStyleSheet()
                hebrew_style = ParagraphStyle(
                    'HebrewStyle', 
                    parent=styles['Normal'], 
                    fontName=PDF_FONT, 
                    fontSize=10, 
                    alignment=TA_RIGHT,
                    wordWrap='RTL'
                )

                elements.append(Paragraph(fix_hebrew("דוח סיכום עסקי כללי"), hebrew_style))
                elements.append(Spacer(1, 10))

                table_data = [[
                    Paragraph(fix_hebrew("רווח נקי"), hebrew_style),
                    Paragraph(fix_hebrew("הוצאות"), hebrew_style),
                    Paragraph(fix_hebrew("חוזה"), hebrew_style),
                    Paragraph(fix_hebrew("שם פרויקט"), hebrew_style)
                ]]
                
                for _, row in df.iterrows():
                    contract_val = row['סכום חוזה (ללא מע"מ)']
                    exp_val = row['סה"כ הוצאות']
                    profit_val = row['רווח נקי']
                    table_data.append([
                        Paragraph(f"{profit_val:,.0f}", hebrew_style),
                        Paragraph(f"{exp_val:,.0f}", hebrew_style),
                        Paragraph(f"{contract_val:,.0f}", hebrew_style),
                        Paragraph(fix_hebrew(str(row["שם פרויקט"])), hebrew_style),
                    ])

                t = Table(table_data)
                t.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                    ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                    ("FONTNAME", (0, 0), (-1, -1), PDF_FONT),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                    ("GRID", (0, 0), (-1, -1), 1, colors.HexColor("#94a3b8")),
                ]))
                elements.append(t)
                doc.build(elements)
                buffer.seek(0)
                return buffer.getvalue()

            general_pdf_bytes = generate_general_pdf(df_summary)
            st.download_button(
                label="📄 הורד דוח סיכום PDF בעברית",
                data=general_pdf_bytes,
                file_name="Business_Summary.pdf",
                mime="application/pdf",
            )

        st.subheader("טבלת ריכוז פרויקטים, מכרזים ויתרות")
        
        display_summary_df = df_summary.copy()
        total_row = pd.DataFrame([{
            "קוד פרויקט": "סה\"כ",
            "שם פרויקט": "סה\"כ כללי",
            "סטטוס": "",
            "תאריך פתיחה": "",
            "תאריך סגירה": "",
            "סכום חוזה (ללא מע\"מ)": tot_contracts,
            "תקבולים בפועל (ללא מע\"מ)": tot_paid,
            "יתרה לתשלום (כולל מע\"מ)": tot_remaining_vat,
            "סה\"כ הוצאות": tot_exp_all,
            "רווח נקי": tot_net_profit,
        }])
        display_summary_df = pd.concat([display_summary_df, total_row], ignore_index=True)
        st.dataframe(display_summary_df, use_container_width=True)

else:
    project_list = df_projects["שם פרויקט"].unique()
    if len(project_list) == 0:
        st.warning("אין פרויקטים במערכת.")
        st.stop()

    selected_project_name = st.selectbox("🎯 בחר פרויקט / מכרז לצפייה וניהול:", project_list)

    current_project = df_projects[df_projects["שם פרויקט"] == selected_project_name].iloc[0]
    project_code = current_project["קוד פרויקט"]
    contract_amount = current_project["סכום חוזה"]
    already_paid = current_project["שולם בפועל"] if "שולם בפועל" in df_projects.columns else 0.0
    current_status = current_project["סטטוס"] if "סטטוס" in df_projects.columns else "בביצוע"
    open_date_val = current_project["תאריך פתיחה"] if "תאריך פתיחה" in df_projects.columns else ""
    close_date_val = str(current_project["תאריך סגירה"]) if "תאריך סגירה" in df_projects.columns and pd.notna(current_project["תאריך סגירה"]) else ""

    project_expenses = df_expenses[df_expenses["קוד פרויקט"] == project_code]
    total_exp = project_expenses["סכום"].sum() if not project_expenses.empty else 0.0
    
    contract_with_vat = contract_amount * (1 + vat_rate)
    paid_with_vat = already_paid * (1 + vat_rate)
    remaining_to_pay = contract_with_vat - paid_with_vat
    net_profit = contract_amount - total_exp

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("סכום חוזה (ללא מע\"מ)", f"{contract_amount:,.0f} ₪")
    col2.metric("יתרה לתשלום (כולל מע\"מ)", f"{remaining_to_pay:,.0f} ₪")
    col3.metric("סה\"כ הוצאות", f"{total_exp:,.0f} ₪")
    col4.metric("רווח נקי", f"{net_profit:,.0f} ₪")

    st.divider()

    with st.expander("⚙️ הגדרות, עריכת שם פרויקט, סטטוסים וניהול"):
        col_upd, col_del = st.columns(2)

        with col_upd:
            with st.form(f"update_payment_form_{project_code}"):
                st.subheader("עדכון פרטי הפרויקט / מכרז ושם פרויקט")
                new_name_input = st.text_input("שם פרויקט מעודכן:", value=str(selected_project_name))
                new_paid_input = st.number_input("עדכן תקבולים בפועל (ללא מע\"מ):", min_value=0.0, value=float(already_paid), step=1000.0)
                status_options = ["מכרז", "בביצוע", "מעוכב", "בבדיקה", "נמסר", "סגור"]
                default_idx = status_options.index(current_status) if current_status in status_options else 0
                new_status = st.selectbox("סטטוס פרויקט:", status_options, index=default_idx)

                update_btn = st.form_submit_button("שמור שינויים")
                if update_btn and new_name_input:
                    updated_close_date = close_date_val
                    if new_status in ["נמסר", "סגור"] and not close_date_val:
                        updated_close_date = datetime.now().strftime("%Y-%m-%d")
                    elif new_status in ["מכרז", "בביצוע", "מעוכב", "בבדיקה"]:
                        updated_close_date = ""

                    df_projects.loc[df_projects["קוד פרויקט"] == project_code, "שם פרויקט"] = new_name_input
                    df_projects.loc[df_projects["קוד פרויקט"] == project_code, "שולם בפועל"] = new_paid_input
                    df_projects.loc[df_projects["קוד פרויקט"] == project_code, "סטטוס"] = new_status
                    df_projects["תאריך סגירה"] = df_projects["תאריך סגירה"].astype(str)
                    df_projects.loc[df_projects["קוד פרויקט"] == project_code, "תאריך סגירה"] = updated_close_date

                    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                        df_projects.to_excel(writer, sheet_name="פרויקטים", index=False)
                        df_expenses.to_excel(writer, sheet_name="הוצאות", index=False)
                    st.success("שם הפרויקט והנתונים עודכנו בהצלחה!")
                    st.rerun()

        with col_del:
            st.warning("מחיקת פרויקט תסיר אותו לצמיתות מהמערכת.")
            if st.button("🗑️ מחק פרויקט זה", key=f"del_proj_{project_code}"):
                df_projects = df_projects[df_projects["קוד פרויקט"] != project_code]
                df_expenses = df_expenses[df_expenses["קוד פרויקט"] != project_code]
                with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                    df_projects.to_excel(writer, sheet_name="פרויקטים", index=False)
                    df_expenses.to_excel(writer, sheet_name="הוצאות", index=False)
                st.success("הפרויקט נמחק בהצלחה!")
                st.rerun()

    st.divider()

    with st.form(f"add_expense_form_{project_code}"):
        st.subheader("➕ הוספת ספק / הוצאה חדשה לפרויקט")
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            new_exp_supplier = st.text_input("שם ספק / תחום הוצאה:")
        with col_e2:
            new_exp_amount = st.number_input("סכום (ללא מע\"מ):", min_value=0.0, step=500.0)
        
        add_exp_btn = st.form_submit_button("הוסף הוצאה")
        if add_exp_btn and new_exp_supplier and new_exp_amount > 0:
            new_exp_row = pd.DataFrame([{
                "קוד פרויקט": project_code,
                "ספק / תחום": new_exp_supplier.strip(),
                "סכום": float(new_exp_amount)
            }])
            df_expenses = pd.concat([df_expenses, new_exp_row], ignore_index=True)
            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                df_projects.to_excel(writer, sheet_name="פרויקטים", index=False)
                df_expenses.to_excel(writer, sheet_name="הוצאות", index=False)
            st.success(f"הספק '{new_exp_supplier}' נוסף בהצלחה!")
            st.rerun()

    st.divider()

    st.subheader(f"📋 פירוט הוצאות וספקים עבור: {selected_project_name}")
    project_expenses = df_expenses[df_expenses["קוד פרויקט"] == project_code]
    if not project_expenses.empty:
        for exp_idx, row in project_expenses.iterrows():
            col_row_info, col_row_del = st.columns([4, 1])
            with col_row_info:
                st.markdown(f"🔹 **{row['ספק / תחום']}** | סכום: {row['סכום']:,.0f} ₪")
            with col_row_del:
                if st.button("🗑️ מחיקה", key=f"del_exp_{exp_idx}"):
                    df_expenses = df_expenses.drop(exp_idx)
                    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                        df_projects.to_excel(writer, sheet_name="פרויקטים", index=False)
                        df_expenses.to_excel(writer, sheet_name="הוצאות", index=False)
                    st.success("ההוצאה נמחקה בהצלחה!")
                    st.rerun()
    else:
        st.info("עדיין אין הוצאות רשומות בפרויקט זה.")

    st.divider()

    st.subheader(f"📥 הפקת דוחות עבור פרויקט: {selected_project_name}")
    col_p_dl1, col_p_dl2 = st.columns(2)

    with col_p_dl1:
        single_proj_df = df_projects[df_projects["קוד פרויקט"] == project_code]
        single_exp_df = project_expenses.copy()
        proj_excel_bytes = to_excel_download({"פרטי פרויקט": single_proj_df, "הוצאות פרויקט": single_exp_df})
        st.download_button(
            label="📥 הורד קובץ Excel לפרויקט (בעברית מלאה)",
            data=proj_excel_bytes,
            file_name=f"Project_{project_code}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col_p_dl2:
        def generate_project_pdf(p_name, p_code, c_amt, paid_amt, rem_pay, tot_ex, prof, exp_df):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            elements = []
            styles = getSampleStyleSheet()
            hebrew_style = ParagraphStyle(
                'HebrewStyle', 
                parent=styles['Normal'], 
                fontName=PDF_FONT, 
                fontSize=11, 
                alignment=TA_RIGHT,
                wordWrap='RTL'
            )

            elements.append(Paragraph(fix_hebrew(f"דוח פרויקט: {p_name}"), hebrew_style))
            elements.append(Spacer(1, 10))

            summary_table_data = [
                [Paragraph(f"{c_amt:,.0f} ₪", hebrew_style), Paragraph(fix_hebrew("סכום חוזה"), hebrew_style)],
                [Paragraph(f"{paid_amt:,.0f} ₪", hebrew_style), Paragraph(fix_hebrew("תקבולים בפועל"), hebrew_style)],
                [Paragraph(f"{rem_pay:,.0f} ₪", hebrew_style), Paragraph(fix_hebrew("יתרה לתשלום (כולל מע\"מ)"), hebrew_style)],
                [Paragraph(f"{tot_ex:,.0f} ₪", hebrew_style), Paragraph(fix_hebrew("סה\"כ הוצאות"), hebrew_style)],
                [Paragraph(f"{prof:,.0f} ₪", hebrew_style), Paragraph(fix_hebrew("רווח נקי"), hebrew_style)],
            ]
            t_sum = Table(summary_table_data)
            t_sum.setStyle(TableStyle([
                ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                ("FONTNAME", (0, 0), (-1, -1), PDF_FONT),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ]))
            elements.append(t_sum)
            elements.append(Spacer(1, 15))

            elements.append(Paragraph(fix_hebrew("פירוט ספקים והוצאות"), hebrew_style))
            exp_table_data = [[
                Paragraph(fix_hebrew("סכום (ללא מע\"מ)"), hebrew_style),
                Paragraph(fix_hebrew("ספק / תחום"), hebrew_style)
            ]]
            if not exp_df.empty:
                for _, r in exp_df.iterrows():
                    exp_table_data.append([
                        Paragraph(f"{r['סכום']:,.0f} ₪", hebrew_style),
                        Paragraph(fix_hebrew(str(r["ספק / תחום"])), hebrew_style)
                    ])
            else:
                exp_table_data.append([
                    Paragraph("-", hebrew_style),
                    Paragraph(fix_hebrew("אין הוצאות רשומות"), hebrew_style)
                ])

            t_exp = Table(exp_table_data)
            t_exp.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f766e")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                ("FONTNAME", (0, 0), (-1, -1), PDF_FONT),
                ("GRID", (0, 0), (-1, -1), 1, colors.HexColor("#0f172a")),
            ]))
            elements.append(t_exp)
            doc.build(elements)
            buffer.seek(0)
            return buffer.getvalue()

        proj_pdf_bytes = generate_project_pdf(selected_project_name, project_code, contract_amount, already_paid, remaining_to_pay, total_exp, net_profit, project_expenses)
        st.download_button(
            label="📄 הורד דוח PDF לפרויקט בעברית",
            data=proj_pdf_bytes,
            file_name=f"Project_{project_code}.pdf",
            mime="application/pdf",
        )
