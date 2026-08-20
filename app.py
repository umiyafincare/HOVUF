import os
import urllib.parse
from datetime import datetime, time
import io
import sqlite3
import pandas as pd
import streamlit as st

# ReportLab for PDF Bill & Reports
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image as PDFImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Company Details
COMPANY_NAME = "HARI OM VISA & UMIYA FINCARE"
COMPANY_ADDRESS = "F-46 VATSALY STATUS, NR. DHAVAL PLAZA, KADI - 384440"
COMPANY_MOBILE = "7698564672 / 9714776364"
COMPANY_TAGLINE = "Visa Consultancy | Insurance & Land Advisor | Property Solution | Daily Accounting"

# Image File Names
LOGO_VISA = "HARI OM.jpg"
LOGO_FINCARE = "UMIYA FIN.jpg"
LOGO_INSURANCE = "HARI OM IL.jpg"
LOGO_PROPERTY = "SHREE UNIYA.jpg"

DB_FILE = "rojmed_ledger.db"
DEFAULT_PIN = "1234"

# Page Settings
st.set_page_config(page_title=COMPANY_NAME, page_icon="💼", layout="wide")

# ----------------- MODERN LIGHT THEME CSS -----------------
st.markdown("""
    <style>
        .stApp {
            background-color: #F8FAFC;
            color: #0F172A;
            font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
        }
        section[data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 1px solid #E2E8F0;
        }
        .stButton>button {
            border-radius: 8px;
            font-weight: 600;
            font-size: 14px;
            padding: 8px 16px;
            border: 1px solid #CBD5E1;
            background-color: #FFFFFF;
            color: #1E293B;
        }
        .stButton>button:hover {
            border-color: #2563EB;
            color: #2563EB;
            background-color: #F1F5F9;
        }
        button[kind="primary"] {
            background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
            color: #FFFFFF !important;
            border: none !important;
        }
        .kpi-card {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            padding: 16px 20px;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
            margin-bottom: 12px;
        }
        .kpi-label {
            font-size: 13px;
            font-weight: 600;
            color: #64748B;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 4px;
        }
        .kpi-value {
            font-size: 24px;
            font-weight: 800;
            color: #0F172A;
            word-wrap: break-word;
            margin-bottom: 4px;
        }
        .kpi-sub {
            font-size: 12px;
            color: #475569;
            font-weight: 500;
        }
    </style>
""", unsafe_allow_html=True)

# ----------------- SQL DATABASE MANAGER -----------------
class SQLManager:
    @staticmethod
    def get_connection():
        conn = sqlite3.connect(DB_FILE, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        return conn

    @staticmethod
    def init_db():
        conn = SQLManager.get_connection()
        c = conn.cursor()
        
        c.execute("""
            CREATE TABLE IF NOT EXISTS Customers (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Created_Date TEXT,
                Customer_Name TEXT,
                Mobile_Number TEXT UNIQUE,
                City_Address TEXT,
                Primary_Service TEXT,
                Notes TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS Invoices_Archive (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Invoice_No TEXT UNIQUE,
                Date TEXT,
                Customer_Name TEXT,
                Mobile_Number TEXT,
                Service_1 TEXT,
                Amount_1 REAL,
                Service_2 TEXT,
                Amount_2 REAL,
                Total_Amount REAL,
                Paid_Amount REAL,
                Pending_Amount REAL,
                Payment_Mode TEXT,
                Remarks TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS Income (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Date TEXT,
                Customer_Person TEXT,
                Work_Details TEXT,
                Amount REAL,
                Payment_Mode TEXT,
                Notes TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS Expense (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Date TEXT,
                Expense_Name TEXT,
                Amount REAL,
                Notes TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS Udhar_Baki (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Date TEXT,
                Customer_Name TEXT,
                Mobile_Number TEXT,
                Service_Details TEXT,
                Total_Amount REAL,
                Paid_Amount REAL,
                Pending_Amount REAL,
                Due_Date TEXT,
                Status TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS Task_Reminder (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                Date TEXT,
                Time TEXT,
                Person_Name TEXT,
                Mobile TEXT,
                Task_Details TEXT,
                Status TEXT
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS Settings (
                Key TEXT PRIMARY KEY,
                Value TEXT,
                Updated_Date TEXT
            )
        """)

        c.execute("INSERT OR IGNORE INTO Settings (Key, Value, Updated_Date) VALUES ('Cash_Opening_Balance', '0.0', ?)", (datetime.now().strftime("%Y-%m-%d"),))
        c.execute("INSERT OR IGNORE INTO Settings (Key, Value, Updated_Date) VALUES ('Bank_Opening_Balance', '0.0', ?)", (datetime.now().strftime("%Y-%m-%d"),))
        c.execute("INSERT OR IGNORE INTO Settings (Key, Value, Updated_Date) VALUES ('Master_PIN', ?, ?)", (DEFAULT_PIN, datetime.now().strftime("%Y-%m-%d")))

        conn.commit()
        conn.close()

    @staticmethod
    def get_df(table_name):
        SQLManager.init_db()
        conn = SQLManager.get_connection()
        try:
            df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
            conn.close()
            if df is not None and not df.empty:
                df.columns = [c.replace('_', ' ').strip() if c != 'ID' else 'ID' for c in df.columns]
                
                # Normalization
                if table_name.lower() == "customers":
                    if "Primary Service" not in df.columns:
                        for alt in ["Primary Service / Purpose", "Primary_Service", "PrimaryService", "Service"]:
                            if alt in df.columns:
                                df["Primary Service"] = df[alt]
                                break
                    if "City Address" not in df.columns:
                        for alt in ["City/Address", "City_Address", "Address", "City"]:
                            if alt in df.columns:
                                df["City Address"] = df[alt]
                                break
            return df
        except Exception:
            conn.close()
            return pd.DataFrame()

    @staticmethod
    def sync_customer(name, phone, service, remarks=""):
        clean_p = str(phone).strip()
        if not clean_p or not name:
            return
        conn = SQLManager.get_connection()
        c = conn.cursor()
        now_dt = datetime.now().strftime("%Y-%m-%d")
        
        c.execute("SELECT ID FROM Customers WHERE Mobile_Number = ?", (clean_p,))
        row = c.fetchone()
        if row:
            c.execute("UPDATE Customers SET Customer_Name = ?, Primary_Service = ? WHERE ID = ?", (name.strip(), service, row['ID']))
        else:
            c.execute("INSERT INTO Customers (Created_Date, Customer_Name, Mobile_Number, City_Address, Primary_Service, Notes) VALUES (?, ?, ?, ?, ?, ?)",
                      (now_dt, name.strip(), clean_p, 'Kadi', service, remarks))
        conn.commit()
        conn.close()

    @staticmethod
    def get_pin():
        conn = SQLManager.get_connection()
        c = conn.cursor()
        c.execute("SELECT Value FROM Settings WHERE Key = 'Master_PIN'")
        row = c.fetchone()
        conn.close()
        return str(row['Value']).strip() if row else DEFAULT_PIN

    @staticmethod
    def set_pin(new_pin):
        conn = SQLManager.get_connection()
        c = conn.cursor()
        c.execute("UPDATE Settings SET Value = ?, Updated_Date = ? WHERE Key = 'Master_PIN'", (str(new_pin).strip(), datetime.now().strftime("%Y-%m-%d")))
        conn.commit()
        conn.close()

    @staticmethod
    def get_opening_balance():
        conn = SQLManager.get_connection()
        c = conn.cursor()
        c.execute("SELECT Value FROM Settings WHERE Key = 'Cash_Opening_Balance'")
        row_c = c.fetchone()
        c.execute("SELECT Value FROM Settings WHERE Key = 'Bank_Opening_Balance'")
        row_b = c.fetchone()
        conn.close()
        cash_op = float(row_c['Value']) if row_c else 0.0
        bank_op = float(row_b['Value']) if row_b else 0.0
        return cash_op, bank_op

    @staticmethod
    def set_opening_balance(cash_op, bank_op):
        conn = SQLManager.get_connection()
        c = conn.cursor()
        dt = datetime.now().strftime("%Y-%m-%d")
        c.execute("UPDATE Settings SET Value = ?, Updated_Date = ? WHERE Key = 'Cash_Opening_Balance'", (str(cash_op), dt))
        c.execute("UPDATE Settings SET Value = ?, Updated_Date = ? WHERE Key = 'Bank_Opening_Balance'", (str(bank_op), dt))
        conn.commit()
        conn.close()

    @staticmethod
    def export_sqlite_to_excel_buffer():
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            for tbl in ["Customers", "Invoices_Archive", "Income", "Expense", "Udhar_Baki", "Task_Reminder", "Settings"]:
                df = SQLManager.get_df(tbl)
                df.to_excel(writer, sheet_name=tbl, index=False)
        buf.seek(0)
        return buf

    @staticmethod
    def import_excel_to_sqlite(uploaded_file):
        try:
            excel_data = pd.read_excel(uploaded_file, sheet_name=None)
            conn = SQLManager.get_connection()
            
            for sheet_name, df in excel_data.items():
                target_table = sheet_name.replace(" ", "_")
                df.columns = [c.replace(" ", "_").replace("/", "_") for c in df.columns]
                
                # Column alias mappings
                if "Primary_Service___Purpose" in df.columns:
                    df = df.rename(columns={"Primary_Service___Purpose": "Primary_Service"})
                if "Customer_Person" not in df.columns and "Customer_Name" in df.columns and target_table == "Income":
                    df = df.rename(columns={"Customer_Name": "Customer_Person"})
                
                if target_table in ["Customers", "Invoices_Archive", "Income", "Expense", "Udhar_Baki", "Task_Reminder", "Settings"]:
                    df.to_sql(target_table, conn, if_exists="replace", index=False)
            
            conn.commit()
            conn.close()
            return True, "Success"
        except Exception as e:
            return False, str(e)

# Initialize SQL Database
SQLManager.init_db()

def render_top_logos():
    with st.container():
        cols = st.columns([1.2, 1.2, 4, 1.2, 1.2])
        with cols[0]:
            if os.path.exists(LOGO_VISA):
                st.image(LOGO_VISA, use_container_width=True)
        with cols[1]:
            if os.path.exists(LOGO_INSURANCE):
                st.image(LOGO_INSURANCE, use_container_width=True)
        with cols[2]:
            st.markdown(f"""
                <div style="text-align: center;">
                    <h2 style="color: #1E3A8A; margin: 0; font-size: 24px; font-weight: 800;">{COMPANY_NAME}</h2>
                    <p style="margin: 3px 0; font-size: 12.5px; color: #475569; font-weight: 500;">📍 {COMPANY_ADDRESS}</p>
                    <p style="margin: 0; font-size: 13px; font-weight: 700; color: #DC2626;">📞 Phone: {COMPANY_MOBILE}</p>
                </div>
            """, unsafe_allow_html=True)
        with cols[3]:
            if os.path.exists(LOGO_FINCARE):
                st.image(LOGO_FINCARE, use_container_width=True)
        with cols[4]:
            if os.path.exists(LOGO_PROPERTY):
                st.image(LOGO_PROPERTY, use_container_width=True)
        st.markdown("<hr style='margin: 12px 0 20px 0; border: 0; height: 1px; background: #E2E8F0;'>", unsafe_allow_html=True)

# ----------------- LOGIN SCREEN -----------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    render_top_logos()
    col_c1, col_c2, col_c3 = st.columns([1, 1.2, 1])
    with col_c2:
        st.markdown("""
            <div style="background: #FFFFFF; padding: 25px; border-radius: 12px; border: 1px solid #E2E8F0; text-align: center;">
                <h3 style="color: #1E3A8A; margin-top: 0;">🔒 Secure SQL Ledger Login</h3>
                <p style="color: #64748B; font-size: 13px;">Powered by High-Performance Relational SQL Database.</p>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        entered_pin = st.text_input("Enter 4-Digit Security PIN:", type="password", max_chars=8)
        
        if st.button("🔓 Unlock & Login", type="primary", use_container_width=True):
            if entered_pin == SQLManager.get_pin():
                st.session_state.logged_in = True
                st.success("Access Granted!")
                st.rerun()
            else:
                st.error("❌ Invalid PIN! Please try again.")
        st.caption(f"Default setup PIN is: `{DEFAULT_PIN}` (Change it in Security settings).")
    st.stop()

render_top_logos()

# ----------------- NAVIGATION MENU -----------------
if "current_page" not in st.session_state:
    st.session_state.current_page = "📊 Dashboard"

st.sidebar.markdown("<h4 style='color:#1E3A8A;'>📌 Navigation Menu</h4>", unsafe_allow_html=True)
menu_items = [
    ("📊 Dashboard", "Business metrics & live summary"),
    ("⏰ Task Reminders", "Calendar & Clock Reminders"),
    ("🧾 Generate Bill / Voucher", "Create, Edit, Print Invoices & Vouchers"),
    ("📋 Due Collections", "Customer Pending Dues & Edit Records"),
    ("📄 Reports & PDF", "Financial Statements & PDF Export"),
    ("🏦 Opening Balance", "Set Starting Balances"),
    ("👥 Customers Directory", "Manage Clients & Broadcasts"),
    ("💰 Income", "View & Manage Income"),
    ("💸 Expenses", "View & Manage Expenses"),
    ("💾 Backup & Restore (Excel / SQL)", "Export / Import Excel & SQL Database"),
    ("⚙️ Security / Change PIN", "Change Master PIN")
]

for label, desc in menu_items:
    is_active = st.session_state.current_page == label
    btn_type = "primary" if is_active else "secondary"
    if st.sidebar.button(label, key=f"nav_{label}", type=btn_type, use_container_width=True):
        st.session_state.current_page = label
        st.rerun()

menu = st.session_state.current_page
st.sidebar.markdown("---")

excel_buf = SQLManager.export_sqlite_to_excel_buffer()
st.sidebar.download_button(
    "📥 Quick Excel Backup", 
    data=excel_buf, 
    file_name=f"SQL_Backup_{datetime.now().strftime('%Y%m%d')}.xlsx", 
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
    use_container_width=True
)

if st.sidebar.button("🔒 Logout System", use_container_width=True):
    st.session_state.logged_in = False
    st.rerun()

SERVICE_OPTIONS = [
    "VISA", "PASSPORT", "PANCARD", "LOAN", "LAND RECORD (7/12 & 8-A)",
    "E-KYC ALL", "FARMER REGISTRATION", "PM-KISAN", "DIGITAL GUJARAT SERVICES",
    "PROVIDENT FUND (PF)", "REVENUE WORK", "MARRIAGE CERTIFICATE", "INSURANCE",
    "AIR TICKET", "OTHER"
]

def search_customer_profile(search_text):
    if not search_text:
        return None, 0.0, []
    clean_q = str(search_text).strip()
    df_c = SQLManager.get_df("Customers")
    df_b = SQLManager.get_df("Udhar_Baki")
    
    matched_cust = None
    if not df_c.empty and "Mobile Number" in df_c:
        m_phone = df_c[df_c["Mobile Number"].astype(str).str.strip() == clean_q]
        if not m_phone.empty:
            matched_cust = m_phone.iloc[0]
        elif "Customer Name" in df_c:
            m_name = df_c[df_c["Customer Name"].astype(str).str.lower() == clean_q.lower()]
            if not m_name.empty:
                matched_cust = m_name.iloc[0]
                
    total_due = 0.0
    due_records = []
    if not df_b.empty and "Pending Amount" in df_b:
        cond = (df_b["Mobile Number"].astype(str).str.strip() == clean_q) if "Mobile Number" in df_b else pd.Series([False]*len(df_b))
        if "Customer Name" in df_b:
            cond = cond | (df_b["Customer Name"].astype(str).str.lower() == clean_q.lower())
        m_due = df_b[cond]
        active_dues = m_due[m_due["Pending Amount"] > 0]
        if not active_dues.empty:
            total_due = float(active_dues["Pending Amount"].sum())
            due_records = active_dues
            
    return matched_cust, total_due, due_records

def generate_invoice_pdf_buffer(bill_no, bill_date, cust_name, cust_phone, service_1, amt_1, service_2, amt_2, total_bill, rec_amt, baki_amt, pay_mode, remarks):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=20, bottomMargin=20)
    elems = []
    styles = getSampleStyleSheet()
    c_title = ParagraphStyle('C1', parent=styles['Heading1'], fontSize=15, leading=17, textColor=colors.HexColor("#1E3A8A"), alignment=1)
    c_sub = ParagraphStyle('C2', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor("#334155"), alignment=1)
    c_inv = ParagraphStyle('C3', parent=styles['Heading2'], fontSize=12, leading=14, textColor=colors.HexColor("#15803D"), alignment=1)
    
    logo_l = PDFImage(LOGO_VISA, width=55, height=55) if os.path.exists(LOGO_VISA) else ""
    logo_r = PDFImage(LOGO_FINCARE, width=80, height=40) if os.path.exists(LOGO_FINCARE) else ""
    
    hdr_table_data = [[logo_l, [Paragraph(f"<b>{COMPANY_NAME}</b>", c_title), Paragraph(f"📍 {COMPANY_ADDRESS}", c_sub), Paragraph(f"<b>📞 Phone:</b> {COMPANY_MOBILE}", c_sub), Paragraph(f"<i>{COMPANY_TAGLINE}</i>", c_sub)], logo_r]]
    t_hdr = Table(hdr_table_data, colWidths=[70, 400, 80])
    t_hdr.setStyle(TableStyle([('ALIGN', (0, 0), (0, -1), 'LEFT'), ('ALIGN', (1, 0), (1, -1), 'CENTER'), ('ALIGN', (2, 0), (2, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    elems.append(t_hdr)
    elems.append(Spacer(1, 5))
    elems.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceAfter=8))
    elems.append(Paragraph("TAX INVOICE / PAYMENT RECEIPT", c_inv))
    elems.append(Spacer(1, 8))
    
    status_text = "PAID" if baki_amt <= 0 else f"PARTIAL (Balance: Rs. {baki_amt:,.2f})"
    meta_data = [
        [f"<b>Invoice No:</b> {bill_no}", f"<b>Date:</b> {bill_date}"],
        [f"<b>Name:</b> {cust_name}", f"<b>Mobile:</b> {cust_phone}"],
        [f"<b>Payment Mode:</b> {pay_mode}", f"<b>Status:</b> {status_text}"]
    ]
    t_meta = Table([[Paragraph(c, styles['Normal']) for c in r] for r in meta_data], colWidths=[275, 275])
    t_meta.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F8FAFC")), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")), ('TOPPADDING', (0, 0), (-1, -1), 4), ('BOTTOMPADDING', (0, 0), (-1, -1), 4)]))
    elems.append(t_meta)
    elems.append(Spacer(1, 12))
    
    items_data = [["Sr.", "Description / Particulars", "Amount (Rs.)"]]
    items_data.append(["1", service_1, f"{float(amt_1):,.2f}"])
    if service_2 and float(amt_2) > 0:
        items_data.append(["2", service_2, f"{float(amt_2):,.2f}"])
    items_data.append(["", "<b>TOTAL AMOUNT</b>", f"<b>Rs. {float(total_bill):,.2f}</b>"])
    if baki_amt > 0:
        items_data.append(["", "<b>PAID AMOUNT</b>", f"Rs. {float(rec_amt):,.2f}"])
        items_data.append(["", "<b>BALANCE DUE</b>", f"<b>Rs. {float(baki_amt):,.2f}</b>"])
    
    t_items = Table([[Paragraph(str(c), styles['Normal']) for c in r] for r in items_data], colWidths=[40, 370, 140])
    t_items.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E3A8A")), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")), ('ALIGN', (-1, 0), (-1, -1), 'RIGHT'), ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#F0FDF4")), ('TOPPADDING', (0, 0), (-1, -1), 5), ('BOTTOMPADDING', (0, 0), (-1, -1), 5)]))
    elems.append(t_items)
    elems.append(Spacer(1, 15))
    
    foot_data = [[f"<b>Note:</b> {remarks}", f"For, <b>{COMPANY_NAME}</b>"], ["", "\n\n___________________________\nAuthorized Signature"]]
    t_foot = Table([[Paragraph(c, styles['Normal']) for c in r] for r in foot_data], colWidths=[310, 240])
    t_foot.setStyle(TableStyle([('ALIGN', (1, 0), (1, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'TOP')]))
    elems.append(t_foot)
    doc.build(elems)
    buf.seek(0)
    return buf

# ----------------- 1. DASHBOARD -----------------
if menu == "📊 Dashboard":
    st.subheader("📊 Business Overview (SQL Database)")
    
    df_rem_all = SQLManager.get_df("Task_Reminder")
    if not df_rem_all.empty and "Status" in df_rem_all:
        pending_tasks = df_rem_all[df_rem_all["Status"] == "Pending"]
        if not pending_tasks.empty:
            st.markdown(f"""
                <div style="background: #FEF2F2; border-left: 5px solid #EF4444; padding: 12px 18px; border-radius: 8px; margin-bottom: 15px;">
                    <h4 style="color: #B91C1C; margin: 0 0 6px 0;">🔔 Active Reminders & Pending Tasks ({len(pending_tasks)})</h4>
                </div>
            """, unsafe_allow_html=True)
            for _, tr in pending_tasks.iterrows():
                t_id = tr["ID"]
                col_t1, col_t2, col_t3, col_t4, col_t5 = st.columns([1.5, 1.5, 2.5, 1.5, 1.5])
                col_t1.write(f"📅 **{tr.get('Date')}**")
                col_t2.write(f"⏰ {tr.get('Time')}")
                col_t3.write(f"📌 **{tr.get('Task Details')}** ({tr.get('Person Name', 'Client')})")
                if pd.notna(tr.get('Mobile')):
                    t_msg = f"Reminder regarding: {tr.get('Task Details')} scheduled on {tr.get('Date')} at {tr.get('Time')}."
                    t_url = f"https://wa.me/91{str(tr.get('Mobile')).strip()}?text={urllib.parse.quote(t_msg)}"
                    col_t4.markdown(f"[📲 WhatsApp]({t_url})")
                else:
                    col_t4.write("-")
                if col_t5.button("✅ Complete", key=f"dash_comp_{t_id}"):
                    conn = SQLManager.get_connection()
                    conn.cursor().execute("UPDATE Task_Reminder SET Status = 'Completed' WHERE ID = ?", (t_id,))
                    conn.commit()
                    conn.close()
                    st.success("Task Completed!")
                    st.rerun()
            st.divider()

    df_inc = SQLManager.get_df("Income")
    df_exp = SQLManager.get_df("Expense")
    df_baki = SQLManager.get_df("Udhar_Baki")
    df_cust = SQLManager.get_df("Customers")
    today_str = datetime.now().strftime("%Y-%m-%d")
    
    today_inc = df_inc[df_inc["Date"] == today_str]["Amount"].sum() if not df_inc.empty and "Date" in df_inc and "Amount" in df_inc else 0.0
    total_inc = df_inc["Amount"].sum() if not df_inc.empty and "Amount" in df_inc else 0.0
    today_exp = df_exp[df_exp["Date"] == today_str]["Amount"].sum() if not df_exp.empty and "Date" in df_exp and "Amount" in df_exp else 0.0
    total_exp = df_exp["Amount"].sum() if not df_exp.empty and "Amount" in df_exp else 0.0
    total_baki = df_baki["Pending Amount"].sum() if not df_baki.empty and "Pending Amount" in df_baki else 0.0
    
    cash_op, bank_op = SQLManager.get_opening_balance()
    tot_op = cash_op + bank_op
    closing_net_balance = tot_op + total_inc - total_exp
    total_cust = len(df_cust) if not df_cust.empty else 0

    k_row1_c1, k_row1_c2, k_row1_c3 = st.columns(3)
    with k_row1_c1:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #2563EB;">
                <div class="kpi-label">🏦 Total Opening Balance</div>
                <div class="kpi-value">₹ {tot_op:,.2f}</div>
                <div class="kpi-sub">Cash: ₹ {cash_op:,.2f} | Bank: ₹ {bank_op:,.2f}</div>
            </div>
        """, unsafe_allow_html=True)
    with k_row1_c2:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #16A34A;">
                <div class="kpi-label">💰 Total Revenue (Income)</div>
                <div class="kpi-value" style="color: #15803D;">₹ {total_inc:,.2f}</div>
                <div class="kpi-sub">Today's Inflow: ₹ {today_inc:,.2f}</div>
            </div>
        """, unsafe_allow_html=True)
    with k_row1_c3:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #DC2626;">
                <div class="kpi-label">💸 Total Expenses</div>
                <div class="kpi-value" style="color: #DC2626;">₹ {total_exp:,.2f}</div>
                <div class="kpi-sub">Today's Outflow: ₹ {today_exp:,.2f}</div>
            </div>
        """, unsafe_allow_html=True)

    k_row2_c1, k_row2_c2, k_row2_c3 = st.columns(3)
    with k_row2_c1:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #0284C7;">
                <div class="kpi-label">💼 Net Closing Balance</div>
                <div class="kpi-value" style="color: #0369A1;">₹ {closing_net_balance:,.2f}</div>
                <div class="kpi-sub">Available Business Capital</div>
            </div>
        """, unsafe_allow_html=True)
    with k_row2_c2:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #D97706;">
                <div class="kpi-label">📋 Total Pending Dues</div>
                <div class="kpi-value" style="color: #B45309;">₹ {total_baki:,.2f}</div>
                <div class="kpi-sub">Uncollected Client Receivables</div>
            </div>
        """, unsafe_allow_html=True)
    with k_row2_c3:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #7C3AED;">
                <div class="kpi-label">👥 Registered Clients</div>
                <div class="kpi-value" style="color: #6D28D9;">{total_cust}</div>
                <div class="kpi-sub">Stored in SQL Database</div>
            </div>
        """, unsafe_allow_html=True)

    st.divider()
    st.subheader("📋 Pending Collections & Itemized WhatsApp Reminders")
    if not df_baki.empty and "Pending Amount" in df_baki:
        pending = df_baki[df_baki["Pending Amount"] > 0]
        if not pending.empty:
            for _, r in pending.iterrows():
                b1, b2, b3, b4, b5, b6 = st.columns([2, 2, 2, 2, 2, 2])
                b1.write(f"**{r.get('Customer Name')}**")
                b2.write(f"📞 {r.get('Mobile Number')}")
                serv_name = str(r.get('Service Details', 'Service'))
                b3.write(f"🏷️ *{serv_name}*")
                b4.write(f"Due: **₹ {r.get('Pending Amount'):,.2f}**")
                b5.write(f"Date: {r.get('Due Date')}")
                msg = f"Hello {r.get('Customer Name')}, payment reminder from {COMPANY_NAME}. An outstanding balance of Rs. {r.get('Pending Amount'):,.2f} is pending for {serv_name}. Contact: {COMPANY_MOBILE}"
                wa_url = f"https://wa.me/91{str(r.get('Mobile Number')).strip()}?text={urllib.parse.quote(msg)}"
                b6.markdown(f"[📲 Send WhatsApp]({wa_url})", unsafe_allow_html=True)
        else:
            st.success("All customer accounts clear!")

# ----------------- 2. TASK REMINDERS -----------------
elif menu == "⏰ Task Reminders":
    st.subheader("⏰ Reminders & Task Management")
    tab_new_task, tab_pending_tasks, tab_completed_tasks = st.tabs(["➕ Schedule New Reminder", "⏳ Active Tasks", "✅ Completed History"])
    
    with tab_new_task:
        with st.form("task_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            tdesc = c1.text_input("Task / Follow-up Details *")
            pname = c2.text_input("Associated Person Name *")
            c3, c4 = st.columns(2)
            rphone = c3.text_input("Mobile Number (10 Digits)")
            rdate = c4.date_input("📅 Reminder Date", datetime.now()).strftime("%Y-%m-%d")
            rtime = st.time_input("⏰ Reminder Time", time(11, 0)).strftime("%I:%M %p")
            if st.form_submit_button("💾 Save Reminder", use_container_width=True):
                if tdesc and pname:
                    conn = SQLManager.get_connection()
                    conn.cursor().execute("INSERT INTO Task_Reminder (Date, Time, Person_Name, Mobile, Task_Details, Status) VALUES (?, ?, ?, ?, ?, 'Pending')",
                                          (rdate, rtime, pname, rphone, tdesc))
                    conn.commit()
                    conn.close()
                    st.success("Reminder Saved to SQL Database!")
                    st.rerun()

    with tab_pending_tasks:
        df_rem = SQLManager.get_df("Task_Reminder")
        if not df_rem.empty:
            pending_list = df_rem[df_rem["Status"] == "Pending"]
            for _, r in pending_list.iterrows():
                r_id = r["ID"]
                c1, c2, c3, c4 = st.columns([2, 2, 3, 2])
                c1.write(f"📅 {r['Date']} | ⏰ {r['Time']}")
                c2.write(f"👤 {r.get('Person Name')}")
                c3.write(f"📌 {r.get('Task Details')}")
                if c4.button("✅ Done", key=f"done_{r_id}"):
                    conn = SQLManager.get_connection()
                    conn.cursor().execute("UPDATE Task_Reminder SET Status = 'Completed' WHERE ID = ?", (r_id,))
                    conn.commit()
                    conn.close()
                    st.rerun()

    with tab_completed_tasks:
        df_rem = SQLManager.get_df("Task_Reminder")
        if not df_rem.empty:
            st.dataframe(df_rem[df_rem["Status"] == "Completed"], use_container_width=True)

# ----------------- 3. INVOICE GENERATION -----------------
elif menu == "🧾 Generate Bill / Voucher":
    st.subheader("🧾 Generate, Edit & Manage Invoices / Vouchers (SQL)")
    bill_type = st.radio("Select Action:", [
        "Customer Invoice (Income)", 
        "✏️ Edit / Delete Invoices (Requires PIN)",
        "🖨️ Re-Print Old Invoice", 
        "Payment Voucher (Expense)", 
        "✏️ Edit / Delete Vouchers (Requires PIN)",
        "Settle Old Pending Due"
    ], horizontal=True)
    
    # --- 1. NEW CUSTOMER INVOICE ---
    if bill_type == "Customer Invoice (Income)":
        st.markdown("### 🔍 STEP 1: Quick Customer Lookup & Live Due Detection")
        
        df_all_cust = SQLManager.get_df("Customers")
        col_s_opt1, col_s_opt2 = st.columns([1.5, 2.5])
        
        selected_from_list = None
        if not df_all_cust.empty:
            serv_col = "Primary Service" if "Primary Service" in df_all_cust else df_all_cust.columns[min(5, len(df_all_cust.columns)-1)]
            cust_quick_list = ["-- Quick Choose Registered Client (Optional) --"] + [
                f"{r.get('Customer Name', '')} ({r.get('Mobile Number', '')}) - {r.get(serv_col, '')}" 
                for _, r in df_all_cust.iterrows()
            ]
            chosen_c = col_s_opt1.selectbox("Search from Directory:", cust_quick_list)
            if chosen_c != "-- Quick Choose Registered Client (Optional) --":
                c_idx = cust_quick_list.index(chosen_c) - 1
                selected_from_list = df_all_cust.iloc[c_idx]

        init_name = str(selected_from_list.get("Customer Name", "")) if selected_from_list is not None else ""
        init_phone = str(selected_from_list.get("Mobile Number", "")) if selected_from_list is not None else ""
        init_service = str(selected_from_list.get("Primary Service", "VISA")) if selected_from_list is not None else "VISA"
        
        col_c1, col_c2 = st.columns(2)
        cust_name = col_c1.text_input("Customer Name *", value=init_name)
        cust_phone = col_c2.text_input("Mobile Number (10 Digits) *", value=init_phone)
        
        search_term = cust_phone if cust_phone else cust_name
        if search_term:
            matched_profile, live_due, due_records = search_customer_profile(search_term)
            
            if matched_profile is not None or live_due > 0:
                p_name = matched_profile.get('Customer Name') if matched_profile is not None else cust_name
                p_phone = matched_profile.get('Mobile Number') if matched_profile is not None else cust_phone
                p_city = matched_profile.get('City Address', 'Kadi') if matched_profile is not None else 'Kadi'
                p_serv = matched_profile.get('Primary Service', 'General') if matched_profile is not None else 'General'
                st.markdown(f"""
                    <div style="background: #F0FDF4; border-left: 5px solid #16A34A; padding: 12px 16px; border-radius: 8px; margin: 10px 0;">
                        <h4 style="color: #15803D; margin: 0;">✅ Client Match Found: {p_name}</h4>
                        <p style="margin: 3px 0 0 0; font-size: 13px; color: #334155;">
                            📞 Mobile: <b>{p_phone}</b> | 
                            📍 City: <b>{p_city}</b> | 
                            💼 Previous Service: <b>{p_serv}</b>
                        </p>
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.info("ℹ️ New Client: Details will be **automatically added to Customer Directory** upon bill generation.")

            if live_due > 0:
                st.markdown(f"""
                    <div style="background: #FEF2F2; border-left: 5px solid #DC2626; padding: 12px 16px; border-radius: 8px; margin: 8px 0 15px 0;">
                        <h4 style="color: #B91C1C; margin: 0;">⚠️ Outstanding Pending Due Alert: ₹ {live_due:,.2f}</h4>
                        <p style="margin: 2px 0 0 0; font-size: 13px; color: #7F1D1D;">This customer has an existing unpaid balance.</p>
                    </div>
                """, unsafe_allow_html=True)
                with st.expander("🔎 View Previous Unpaid Dues Breakdown"):
                    st.dataframe(due_records[["Date", "Service Details", "Total Amount", "Paid Amount", "Pending Amount", "Due Date"]], use_container_width=True)
            else:
                st.success("✨ Outstanding Account Status: Clear (No Pending Dues).")

        st.markdown("---")
        st.markdown("### 🧾 STEP 2: Bill & Service Particulars")
        
        c3, c4 = st.columns(2)
        bill_no = c3.text_input("Invoice No.", f"INV-{datetime.now().strftime('%Y%m%d%H%M')}")
        bill_date = c4.date_input("Invoice Date", datetime.now()).strftime("%Y-%m-%d")
        
        c5, c6 = st.columns(2)
        default_s_idx = SERVICE_OPTIONS.index(init_service) if init_service in SERVICE_OPTIONS else 0
        s1_sel = c5.selectbox("Select Service 1 *", SERVICE_OPTIONS, index=default_s_idx)
        s1 = c5.text_input("Custom Service Name *") if s1_sel == "OTHER" else s1_sel
        amt1 = c6.number_input("Amount 1 (₹) *", min_value=0.0, step=100.0)
        
        c7, c8 = st.columns(2)
        s2_sel = c7.selectbox("Service 2 (Optional)", ["None"] + SERVICE_OPTIONS)
        s2 = c7.text_input("Custom Service 2") if s2_sel == "OTHER" else ("" if s2_sel == "None" else s2_sel)
        amt2 = c8.number_input("Amount 2 (₹)", min_value=0.0, step=100.0)
        
        total_bill = amt1 + amt2
        st.write(f"### **Total Current Bill: ₹ {total_bill:,.2f}**")
        
        cp1, cp2, cp3 = st.columns(3)
        pay_mode = cp1.selectbox("Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque", "Pending / Due"])
        rec_amt = cp2.number_input("Received Amount (₹)", min_value=0.0, max_value=float(total_bill), value=float(total_bill) if pay_mode != "Pending / Due" else 0.0, step=100.0)
        due_date = cp3.date_input("Due Date (If balance pending)", datetime.now()).strftime("%Y-%m-%d")
        baki_amt = total_bill - rec_amt
        item_desc = s1 + (f" + {s2}" if s2 else "")
        
        if baki_amt > 0:
            st.warning(f"⚠️ Balance Due on this Bill: ₹ {baki_amt:,.2f} for '{item_desc}'")
            
        remarks = st.text_input("Remarks / Notes", "Thank you for choosing our services!")

        if st.button("💾 Generate Bill & Save to SQL", type="primary", use_container_width=True):
            if cust_name and total_bill > 0 and s1:
                SQLManager.sync_customer(cust_name, cust_phone, s1, remarks)
                
                conn = SQLManager.get_connection()
                c = conn.cursor()
                c.execute("""
                    INSERT INTO Invoices_Archive 
                    (Invoice_No, Date, Customer_Name, Mobile_Number, Service_1, Amount_1, Service_2, Amount_2, Total_Amount, Paid_Amount, Pending_Amount, Payment_Mode, Remarks)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (bill_no, bill_date, cust_name.strip(), str(cust_phone).strip(), s1, amt1, s2, amt2, total_bill, rec_amt, baki_amt, pay_mode, remarks))
                
                if rec_amt > 0:
                    c.execute("INSERT INTO Income (Date, Customer_Person, Work_Details, Amount, Payment_Mode, Notes) VALUES (?, ?, ?, ?, ?, ?)",
                              (bill_date, cust_name.strip(), f"Bill #{bill_no}: {item_desc}", rec_amt, pay_mode, f"Mob: {cust_phone}"))
                
                if baki_amt > 0:
                    c.execute("INSERT INTO Udhar_Baki (Date, Customer_Name, Mobile_Number, Service_Details, Total_Amount, Paid_Amount, Pending_Amount, Due_Date, Status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Pending')",
                              (bill_date, cust_name.strip(), str(cust_phone).strip(), item_desc, total_bill, rec_amt, baki_amt, due_date))
                
                conn.commit()
                conn.close()
                
                pdf_data = generate_invoice_pdf_buffer(bill_no, bill_date, cust_name, cust_phone, s1, amt1, s2, amt2, total_bill, rec_amt, baki_amt, pay_mode, remarks)
                col_dwn, col_wa = st.columns(2)
                col_dwn.download_button("📥 Download PDF Invoice", data=pdf_data, file_name=f"Invoice_{cust_name}_{bill_no}.pdf", mime="application/pdf", type="primary", use_container_width=True)
                
                wa_msg = f"🧾 *TAX INVOICE*\n🏢 *{COMPANY_NAME}*\n📄 *Invoice No:* {bill_no}\n👤 *Customer:* {cust_name}\n💼 *Service:* {item_desc}\n💰 *Total:* Rs. {total_bill:,.2f}\n✅ *Paid:* Rs. {rec_amt:,.2f}\n"
                if baki_amt > 0:
                    wa_msg += f"⚠️ *Pending Due:* Rs. {baki_amt:,.2f} (Due: {due_date})\n"
                wa_msg += f"📞 {COMPANY_MOBILE}\n🙏 *Thank you for your business!*"
                
                wa_url = f"https://wa.me/91{str(cust_phone).strip()}?text={urllib.parse.quote(wa_msg)}"
                col_wa.markdown(f'<a href="{wa_url}" target="_blank"><button style="width:100%; height:45px; background-color:#25D366; color:white; border:none; border-radius:8px; font-weight:bold; cursor:pointer;">📲 Send Invoice via WhatsApp</button></a>', unsafe_allow_html=True)
                st.success("✅ Bill Created & Saved to SQL Database!")
            else:
                st.error("Please enter customer name, valid service, and bill amount.")

    # --- 2. EDIT / DELETE GENERATED INVOICES ---
    elif bill_type == "✏️ Edit / Delete Invoices (Requires PIN)":
        st.markdown("### 🔐 Modify or Delete Existing Invoice Record (SQL)")
        df_arch = SQLManager.get_df("Invoices_Archive")
        
        if not df_arch.empty and "Invoice No" in df_arch:
            sel_inv_options = [f"{r['Invoice No']} - {r['Customer Name']} ({r['Date']}) | Total: ₹{r['Total Amount']}" for _, r in df_arch.iterrows()]
            chosen_inv_str = st.selectbox("Select Invoice to Modify / Delete:", sel_inv_options, key="edit_inv_select")
            
            if chosen_inv_str:
                sel_inv_no = chosen_inv_str.split(" - ")[0]
                inv_row = df_arch[df_arch["Invoice No"] == sel_inv_no].iloc[0]
                
                with st.expander(f"📝 Edit Invoice #{sel_inv_no} Details", expanded=True):
                    ed_c1, ed_c2 = st.columns(2)
                    up_date = ed_c1.text_input("Invoice Date", str(inv_row.get("Date", "")))
                    up_cname = ed_c2.text_input("Customer Name", str(inv_row.get("Customer Name", "")))
                    
                    ed_c3, ed_c4 = st.columns(2)
                    up_cphone = ed_c3.text_input("Mobile Number", str(inv_row.get("Mobile Number", "")))
                    up_s1 = ed_c4.text_input("Service 1", str(inv_row.get("Service 1", "")))
                    
                    ed_c5, ed_c6 = st.columns(2)
                    up_amt1 = ed_c5.number_input("Amount 1 (₹)", value=float(inv_row.get("Amount 1", 0.0)), step=100.0)
                    up_s2 = ed_c6.text_input("Service 2", str(inv_row.get("Service 2", "")) if pd.notna(inv_row.get("Service 2")) else "")
                    
                    ed_c7, ed_c8 = st.columns(2)
                    up_amt2 = ed_c7.number_input("Amount 2 (₹)", value=float(inv_row.get("Amount 2", 0.0)) if pd.notna(inv_row.get("Amount 2")) else 0.0, step=100.0)
                    up_mode = ed_c8.selectbox("Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque", "Pending / Due"], index=["Cash", "UPI / GPay", "Bank Transfer", "Cheque", "Pending / Due"].index(inv_row.get("Payment Mode", "Cash")) if inv_row.get("Payment Mode") in ["Cash", "UPI / GPay", "Bank Transfer", "Cheque", "Pending / Due"] else 0)
                    
                    up_tot = up_amt1 + up_amt2
                    ed_c9, ed_c10 = st.columns(2)
                    up_rec = ed_c9.number_input("Paid Amount (₹)", value=float(inv_row.get("Paid Amount", up_tot)), max_value=float(up_tot), step=100.0)
                    up_baki = up_tot - up_rec
                    ed_c10.write(f"**Calculated Balance Due:** ₹ {up_baki:,.2f}")
                    
                    up_remarks = st.text_input("Remarks", str(inv_row.get("Remarks", "")) if pd.notna(inv_row.get("Remarks")) else "")
                    
                    st.markdown("🔒 **Security Authorization:**")
                    inv_auth_pin = st.text_input("Enter 4-Digit Security PIN to Authorize:", type="password", key=f"inv_pin_{sel_inv_no}")
                    
                    btn_up_col, btn_del_col = st.columns(2)
                    if btn_up_col.button("🔄 Update Invoice Record", key=f"up_btn_{sel_inv_no}", use_container_width=True):
                        if inv_auth_pin == SQLManager.get_pin():
                            conn = SQLManager.get_connection()
                            c = conn.cursor()
                            c.execute("""
                                UPDATE Invoices_Archive SET 
                                Date = ?, Customer_Name = ?, Mobile_Number = ?, Service_1 = ?, Amount_1 = ?, 
                                Service_2 = ?, Amount_2 = ?, Total_Amount = ?, Paid_Amount = ?, Pending_Amount = ?, 
                                Payment_Mode = ?, Remarks = ?
                                WHERE Invoice_No = ?
                            """, (up_date, up_cname, str(up_cphone).strip(), up_s1, up_amt1, up_s2, up_amt2, up_tot, up_rec, up_baki, up_mode, up_remarks, sel_inv_no))
                            conn.commit()
                            conn.close()
                            st.success(f"✅ Invoice #{sel_inv_no} updated successfully in SQL!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
                            
                    if btn_del_col.button("🗑️ Delete Invoice Record", key=f"del_btn_{sel_inv_no}", type="primary", use_container_width=True):
                        if inv_auth_pin == SQLManager.get_pin():
                            conn = SQLManager.get_connection()
                            c = conn.cursor()
                            c.execute("DELETE FROM Invoices_Archive WHERE Invoice_No = ?", (sel_inv_no,))
                            conn.commit()
                            conn.close()
                            st.warning(f"🗑️ Invoice #{sel_inv_no} deleted from SQL Database!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
        else:
            st.info("No generated invoice records found to edit.")

    # --- 3. RE-PRINT OLD INVOICES ---
    elif bill_type == "🖨️ Re-Print Old Invoice":
        df_arch = SQLManager.get_df("Invoices_Archive")
        if not df_arch.empty:
            sel_inv = st.selectbox("Select Invoice:", [f"{r['Invoice No']} - {r['Customer Name']} ({r['Date']})" for _, r in df_arch.iterrows()])
            sel_no = sel_inv.split(" - ")[0]
            r = df_arch[df_arch["Invoice No"] == sel_no].iloc[0]
            re_pdf = generate_invoice_pdf_buffer(str(r["Invoice No"]), str(r["Date"]), str(r["Customer Name"]), str(r["Mobile Number"]), str(r.get("Service 1", "")), float(r.get("Amount 1", 0)), str(r.get("Service 2", "")), float(r.get("Amount 2", 0)), float(r["Total Amount"]), float(r.get("Paid Amount", 0)), float(r.get("Pending Amount", 0)), str(r.get("Payment Mode", "")), str(r.get("Remarks", "")))
            st.download_button("🖨️ Re-Download PDF", data=re_pdf, file_name=f"Invoice_{sel_no}.pdf", mime="application/pdf", type="primary", use_container_width=True)
        else:
            st.info("No invoices found.")

    # --- 4. EXPENSE VOUCHER ---
    elif bill_type == "Payment Voucher (Expense)":
        c1, c2 = st.columns(2)
        v_no = c1.text_input("Voucher No.", f"VOU-{datetime.now().strftime('%Y%m%d%H%M')}")
        v_date = c2.date_input("Date", datetime.now()).strftime("%Y-%m-%d")
        p_name = c1.text_input("Paid To *")
        p_amt = c2.number_input("Amount (₹) *", min_value=0.0, step=50.0)
        p_mode = c1.selectbox("Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"])
        p_desc = c2.text_input("Expense Purpose *")
        if st.button("💾 Save Expense to SQL", type="primary", use_container_width=True):
            if p_name and p_amt > 0:
                conn = SQLManager.get_connection()
                conn.cursor().execute("INSERT INTO Expense (Date, Expense_Name, Amount, Notes) VALUES (?, ?, ?, ?)",
                                      (v_date, f"{p_name} ({p_desc})", p_amt, f"VOU #{v_no} | {p_mode}"))
                conn.commit()
                conn.close()
                st.success("Expense Recorded in SQL Database!")

    # --- 5. EDIT / DELETE EXPENSE VOUCHERS ---
    elif bill_type == "✏️ Edit / Delete Vouchers (Requires PIN)":
        st.markdown("### 🔐 Modify or Delete Payment Voucher Record")
        df_exp = SQLManager.get_df("Expense")
        
        if not df_exp.empty:
            sel_exp_options = [f"ID #{r['ID']} - {r.get('Expense Name', '')} ({r.get('Date', '')}) | ₹{r.get('Amount', 0)}" for _, r in df_exp.iterrows()]
            chosen_exp_str = st.selectbox("Select Voucher / Expense to Modify:", sel_exp_options, key="edit_exp_select")
            
            if chosen_exp_str:
                sel_exp_id = int(chosen_exp_str.split(" ")[1].replace("#", ""))
                exp_row = df_exp[df_exp["ID"] == sel_exp_id].iloc[0]
                
                with st.expander(f"📝 Edit Expense Voucher #{sel_exp_id}", expanded=True):
                    up_ed1, up_ed2 = st.columns(2)
                    up_e_date = up_ed1.text_input("Date", str(exp_row.get("Date", "")), key=f"e_dt_{sel_exp_id}")
                    up_e_name = up_ed2.text_input("Expense Description / Paid To", str(exp_row.get("Expense Name", "")), key=f"e_nm_{sel_exp_id}")
                    
                    up_ed3, up_ed4 = st.columns(2)
                    up_e_amt = up_ed3.number_input("Amount (₹)", value=float(exp_row.get("Amount", 0.0)), step=50.0, key=f"e_am_{sel_exp_id}")
                    up_e_notes = up_ed4.text_input("Notes", str(exp_row.get("Notes", "")) if pd.notna(exp_row.get("Notes")) else "", key=f"e_nt_{sel_exp_id}")
                    
                    st.markdown("🔒 **Security Authorization:**")
                    exp_auth_pin = st.text_input("Enter 4-Digit Security PIN to Authorize:", type="password", key=f"exp_pin_{sel_exp_id}")
                    
                    eb_up_col, eb_del_col = st.columns(2)
                    if eb_up_col.button("🔄 Update Voucher", key=f"up_exp_btn_{sel_exp_id}", use_container_width=True):
                        if exp_auth_pin == SQLManager.get_pin():
                            conn = SQLManager.get_connection()
                            conn.cursor().execute("UPDATE Expense SET Date = ?, Expense_Name = ?, Amount = ?, Notes = ? WHERE ID = ?",
                                                  (up_e_date, up_e_name, up_e_amt, up_e_notes, sel_exp_id))
                            conn.commit()
                            conn.close()
                            st.success(f"✅ Voucher #{sel_exp_id} updated in SQL!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
                            
                    if eb_del_col.button("🗑️ Delete Voucher", key=f"del_exp_btn_{sel_exp_id}", type="primary", use_container_width=True):
                        if exp_auth_pin == SQLManager.get_pin():
                            conn = SQLManager.get_connection()
                            conn.cursor().execute("DELETE FROM Expense WHERE ID = ?", (sel_exp_id,))
                            conn.commit()
                            conn.close()
                            st.warning(f"🗑️ Voucher #{sel_exp_id} deleted from SQL!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
        else:
            st.info("No expense vouchers found to edit.")

    # --- 6. SETTLE OLD PENDING DUE ---
    elif bill_type == "Settle Old Pending Due":
        df_baki = SQLManager.get_df("Udhar_Baki")
        if not df_baki.empty:
            pending = df_baki[df_baki["Pending Amount"] > 0]
            if not pending.empty:
                sel_acc = st.selectbox("Select Due Account:", [f"ID #{r['ID']} - {r['Customer Name']} | Pending: ₹{r['Pending Amount']}" for _, r in pending.iterrows()])
                sel_id = int(sel_acc.split(" ")[1].replace("#", ""))
                
                # Fetch row directly to prevent KeyError
                conn = SQLManager.get_connection()
                c = conn.cursor()
                c.execute("SELECT * FROM Udhar_Baki WHERE ID = ?", (sel_id,))
                r = c.fetchone()
                conn.close()
                
                c_name = str(r['Customer_Name']) if 'Customer_Name' in r.keys() else str(r['Customer Name'])
                c_serv = str(r['Service_Details']) if 'Service_Details' in r.keys() else str(r['Service Details'])
                curr_pend = float(r['Pending_Amount']) if 'Pending_Amount' in r.keys() else float(r['Pending Amount'])
                curr_paid = float(r['Paid_Amount']) if 'Paid_Amount' in r.keys() else float(r['Paid Amount'])
                
                s_amt = st.number_input("Payment Received Now (₹) *", min_value=0.0, max_value=curr_pend, value=curr_pend, step=100.0)
                s_mode = st.selectbox("Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"])
                if st.button("💳 Settle Balance", type="primary", use_container_width=True):
                    new_paid = curr_paid + s_amt
                    new_pending = curr_pend - s_amt
                    new_stat = "Cleared" if new_pending <= 0 else "Pending"
                    
                    conn = SQLManager.get_connection()
                    c = conn.cursor()
                    c.execute("UPDATE Udhar_Baki SET Paid_Amount = ?, Pending_Amount = ?, Status = ? WHERE ID = ?",
                              (new_paid, new_pending, new_stat, sel_id))
                    c.execute("INSERT INTO Income (Date, Customer_Person, Work_Details, Amount, Payment_Mode, Notes) VALUES (?, ?, ?, ?, ?, ?)",
                              (datetime.now().strftime("%Y-%m-%d"), c_name, f"Due Settlement ({c_serv})", s_amt, s_mode, f"Due Rec #{sel_id}"))
                    conn.commit()
                    conn.close()
                    st.success("Due Settled in SQL!")
                    st.rerun()

# ----------------- 4. DUE COLLECTIONS (VIEW, SETTLE, EDIT & DELETE WITH PIN) -----------------
elif menu == "📋 Due Collections":
    st.subheader("📋 Due Collections & Credit Ledger Management (SQL)")
    
    tab_due_view, tab_due_edit = st.tabs(["📋 Active Due Receivables", "🔐 Edit / Modify Due Record (Requires PIN)"])
    
    with tab_due_view:
        df_b = SQLManager.get_df("Udhar_Baki")
        if not df_b.empty:
            st.dataframe(df_b, use_container_width=True)
        else:
            st.info("No due records found.")
            
    with tab_due_edit:
        st.markdown("##### 🔐 Modify or Delete Customer Due Record")
        df_b = SQLManager.get_df("Udhar_Baki")
        if not df_b.empty:
            sel_due_opts = [f"ID #{r['ID']} - {r.get('Customer Name', '')} ({r.get('Mobile Number', '')}) | Pending: ₹{r.get('Pending Amount', 0)}" for _, r in df_b.iterrows()]
            chosen_due_str = st.selectbox("Select Due Record to Edit / Delete:", sel_due_opts, key="due_mod_select")
            
            if chosen_due_str:
                sel_due_id = int(chosen_due_str.split(" ")[1].replace("#", ""))
                due_r = df_b[df_b["ID"] == sel_due_id].iloc[0]
                
                with st.expander(f"📝 Edit Due Record #{sel_due_id} - {due_r.get('Customer Name', '')}", expanded=True):
                    dc1, dc2 = st.columns(2)
                    up_d_date = dc1.text_input("Entry Date", str(due_r.get("Date", "")), key=f"d_dt_{sel_due_id}")
                    up_d_name = dc2.text_input("Customer Name", str(due_r.get("Customer Name", "")), key=f"d_nm_{sel_due_id}")
                    
                    dc3, dc4 = st.columns(2)
                    up_d_phone = dc3.text_input("Mobile Number", str(due_r.get("Mobile Number", "")), key=f"d_ph_{sel_due_id}")
                    up_d_serv = dc4.text_input("Service Details", str(due_r.get("Service Details", "Service")), key=f"d_sv_{sel_due_id}")
                    
                    dc5, dc6 = st.columns(2)
                    up_d_tot = dc5.number_input("Total Amount (₹)", value=float(due_r.get("Total Amount", 0.0)), step=100.0, key=f"d_tot_{sel_due_id}")
                    up_d_paid = dc6.number_input("Paid Amount (₹)", value=float(due_r.get("Paid Amount", 0.0)), max_value=float(up_d_tot), step=100.0, key=f"d_pd_{sel_due_id}")
                    
                    up_d_pending = up_d_tot - up_d_paid
                    st.write(f"**Recalculated Pending Amount:** ₹ {up_d_pending:,.2f}")
                    
                    dc7, dc8 = st.columns(2)
                    up_d_due_dt = dc7.text_input("Due Date", str(due_r.get("Due Date", "")), key=f"d_ddt_{sel_due_id}")
                    up_d_stat = dc8.selectbox("Status", ["Pending", "Cleared"], index=0 if up_d_pending > 0 else 1, key=f"d_st_{sel_due_id}")
                    
                    st.markdown("🔒 **Security Authorization:**")
                    due_auth_pin = st.text_input("Enter 4-Digit Security PIN to Authorize:", type="password", key=f"due_pin_{sel_due_id}")
                    
                    db_up_c, db_del_c = st.columns(2)
                    if db_up_c.button("🔄 Update Due Record", key=f"btn_up_due_{sel_due_id}", use_container_width=True):
                        if due_auth_pin == SQLManager.get_pin():
                            conn = SQLManager.get_connection()
                            conn.cursor().execute("""
                                UPDATE Udhar_Baki SET 
                                Date = ?, Customer_Name = ?, Mobile_Number = ?, Service_Details = ?, 
                                Total_Amount = ?, Paid_Amount = ?, Pending_Amount = ?, Due_Date = ?, Status = ?
                                WHERE ID = ?
                            """, (up_d_date, up_d_name, str(up_d_phone).strip(), up_d_serv, up_d_tot, up_d_paid, up_d_pending, up_d_due_dt, up_d_stat, sel_due_id))
                            conn.commit()
                            conn.close()
                            st.success(f"✅ Due Record #{sel_due_id} updated successfully in SQL!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
                            
                    if db_del_c.button("🗑️ Delete Due Record", key=f"btn_del_due_{sel_due_id}", type="primary", use_container_width=True):
                        if due_auth_pin == SQLManager.get_pin():
                            conn = SQLManager.get_connection()
                            conn.cursor().execute("DELETE FROM Udhar_Baki WHERE ID = ?", (sel_due_id,))
                            conn.commit()
                            conn.close()
                            st.warning(f"🗑️ Due Record #{sel_due_id} deleted from SQL Database!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
        else:
            st.info("No due records found to edit.")

# ----------------- 5. REPORTS & PDF (LANDSCAPE A4 LAYOUT) -----------------
elif menu == "📄 Reports & PDF":
    st.subheader("📄 Financial Reports & Statements")
    c1, c2 = st.columns(2)
    d_from = c1.date_input("From Date", datetime.now().replace(day=1)).strftime("%Y-%m-%d")
    d_to = c2.date_input("To Date", datetime.now()).strftime("%Y-%m-%d")
    
    df_i = SQLManager.get_df("Income")
    df_e = SQLManager.get_df("Expense")
    df_b = SQLManager.get_df("Udhar_Baki")
    
    f_i = df_i[(df_i["Date"] >= d_from) & (df_i["Date"] <= d_to)] if not df_i.empty and "Date" in df_i else pd.DataFrame()
    f_e = df_e[(df_e["Date"] >= d_from) & (df_e["Date"] <= d_to)] if not df_e.empty and "Date" in df_e else pd.DataFrame()
    f_b = df_b[(df_b["Date"] >= d_from) & (df_b["Date"] <= d_to)] if not df_b.empty and "Date" in df_b else (df_b if not df_b.empty else pd.DataFrame())
    
    t_i = f_i["Amount"].sum() if not f_i.empty and "Amount" in f_i else 0.0
    t_e = f_e["Amount"].sum() if not f_e.empty and "Amount" in f_e else 0.0
    cash_op, bank_op = SQLManager.get_opening_balance()
    tot_op = cash_op + bank_op
    closing_bal = tot_op + t_i - t_e
    
    st.info(f"**Period:** {d_from} to {d_to} | **Revenue:** ₹{t_i:,.2f} | **Expenses:** ₹{t_e:,.2f} | **Closing Balance:** ₹{closing_bal:,.2f}")
    
    def generate_statement_pdf_landscape(period_from, period_to, df_inc, df_exp, df_due, op_bal, tot_rev, tot_exp, cl_bal):
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=25, leftMargin=25, topMargin=20, bottomMargin=20)
        elems = []
        styles = getSampleStyleSheet()
        
        c_title = ParagraphStyle('T1', parent=styles['Heading1'], fontSize=16, leading=18, textColor=colors.HexColor("#1E3A8A"), alignment=1)
        c_sub = ParagraphStyle('T2', parent=styles['Normal'], fontSize=9, leading=11, textColor=colors.HexColor("#475569"), alignment=1)
        tbl_text = ParagraphStyle('TT', parent=styles['Normal'], fontSize=8.5, leading=10.5, textColor=colors.HexColor("#0F172A"))
        tbl_hdr = ParagraphStyle('TH', parent=styles['Normal'], fontSize=9, leading=11, textColor=colors.white, fontName="Helvetica-Bold")
        
        logo_l = PDFImage(LOGO_VISA, width=50, height=50) if os.path.exists(LOGO_VISA) else ""
        logo_r = PDFImage(LOGO_FINCARE, width=75, height=38) if os.path.exists(LOGO_FINCARE) else ""
            
        hdr_table_data = [[
            logo_l,
            [
                Paragraph(f"<b>{COMPANY_NAME}</b>", c_title),
                Paragraph(f"📍 {COMPANY_ADDRESS} | 📞 Phone: {COMPANY_MOBILE}", c_sub),
                Paragraph(f"<b>FINANCIAL STATEMENT & BUSINESS LEDGER ({period_from} to {period_to})</b>", ParagraphStyle('T3', parent=styles['Heading2'], fontSize=11, leading=13, alignment=1))
            ],
            logo_r
        ]]
        t_hdr = Table(hdr_table_data, colWidths=[70, 650, 70])
        t_hdr.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elems.append(t_hdr)
        elems.append(Spacer(1, 6))
        elems.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceAfter=8))

        sum_tbl = Table([
            [
                Paragraph("<b>Total Opening Balance:</b>", tbl_text), f"Rs. {op_bal:,.2f}",
                Paragraph("<b>Total Revenue (Income):</b>", tbl_text), f"Rs. {tot_rev:,.2f}"
            ],
            [
                Paragraph("<b>Total Expenses:</b>", tbl_text), f"Rs. {tot_exp:,.2f}",
                Paragraph("<b>Net Closing Balance:</b>", tbl_text), f"Rs. {cl_bal:,.2f}"
            ]
        ], colWidths=[180, 215, 180, 215])
        sum_tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elems.append(sum_tbl)
        elems.append(Spacer(1, 10))

        if not df_inc.empty:
            elems.append(Paragraph("<b>💰 REVENUE / INCOME BREAKDOWN:</b>", styles['Heading3']))
            i_rows = [[
                Paragraph("Date", tbl_hdr), 
                Paragraph("Customer / Client Name", tbl_hdr), 
                Paragraph("Work & Service Details", tbl_hdr), 
                Paragraph("Mode", tbl_hdr), 
                Paragraph("Amount (Rs.)", tbl_hdr)
            ]]
            for _, r in df_inc.iterrows():
                i_rows.append([
                    Paragraph(str(r.get("Date", "-")), tbl_text),
                    Paragraph(str(r.get("Customer Person", "-")), tbl_text),
                    Paragraph(str(r.get("Work Details", "-")), tbl_text),
                    Paragraph(str(r.get("Payment Mode", "-")), tbl_text),
                    Paragraph(f"<b>{float(r.get('Amount', 0)):,.2f}</b>", tbl_text)
                ])
            t1 = Table(i_rows, colWidths=[80, 220, 310, 80, 100])
            t1.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ('ALIGN', (-1, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('TOPPADDING', (0, 0), (-1, -1), 3.5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3.5),
            ]))
            elems.append(t1)
            elems.append(Spacer(1, 10))

        if not df_exp.empty:
            elems.append(Paragraph("<b>💸 EXPENSE BREAKDOWN:</b>", styles['Heading3']))
            e_rows = [[
                Paragraph("Date", tbl_hdr), 
                Paragraph("Expense Particulars / Description", tbl_hdr), 
                Paragraph("Notes / Remarks", tbl_hdr), 
                Paragraph("Amount (Rs.)", tbl_hdr)
            ]]
            for _, r in df_exp.iterrows():
                e_rows.append([
                    Paragraph(str(r.get("Date", "-")), tbl_text),
                    Paragraph(str(r.get("Expense Name", "-")), tbl_text),
                    Paragraph(str(r.get("Notes", "-")), tbl_text),
                    Paragraph(f"<b>{float(r.get('Amount', 0)):,.2f}</b>", tbl_text)
                ])
            t2 = Table(e_rows, colWidths=[80, 380, 230, 100])
            t2.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#DC2626")),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ('ALIGN', (-1, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('TOPPADDING', (0, 0), (-1, -1), 3.5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3.5),
            ]))
            elems.append(t2)
            elems.append(Spacer(1, 10))

        if not df_due.empty:
            active_dues = df_due[df_due["Pending Amount"] > 0] if "Pending Amount" in df_due else pd.DataFrame()
            if not active_dues.empty:
                elems.append(Paragraph("<b>📋 OUTSTANDING PENDING DUES (RECEIVABLES):</b>", styles['Heading3']))
                d_rows = [[
                    Paragraph("Date", tbl_hdr), 
                    Paragraph("Customer Name", tbl_hdr), 
                    Paragraph("Mobile", tbl_hdr), 
                    Paragraph("Service Details", tbl_hdr), 
                    Paragraph("Total (Rs.)", tbl_hdr), 
                    Paragraph("Paid (Rs.)", tbl_hdr), 
                    Paragraph("Balance (Rs.)", tbl_hdr), 
                    Paragraph("Due Date", tbl_hdr)
                ]]
                for _, r in active_dues.iterrows():
                    d_rows.append([
                        Paragraph(str(r.get("Date", "-")), tbl_text),
                        Paragraph(str(r.get("Customer Name", "-")), tbl_text),
                        Paragraph(str(r.get("Mobile Number", "-")), tbl_text),
                        Paragraph(str(r.get("Service Details", "Service")), tbl_text),
                        Paragraph(f"{float(r.get('Total Amount', 0)):,.2f}", tbl_text),
                        Paragraph(f"{float(r.get('Paid Amount', 0)):,.2f}", tbl_text),
                        Paragraph(f"<b>{float(r.get('Pending Amount', 0)):,.2f}</b>", tbl_text),
                        Paragraph(str(r.get("Due Date", "-")), tbl_text)
                    ])
                t3 = Table(d_rows, colWidths=[70, 160, 95, 175, 75, 75, 75, 65])
                t3.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#D97706")),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                    ('ALIGN', (4, 0), (6, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('TOPPADDING', (0, 0), (-1, -1), 3.5),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3.5),
                ]))
                elems.append(t3)

        doc.build(elems)
        buf.seek(0)
        return buf

    stat_pdf = generate_statement_pdf_landscape(d_from, d_to, f_i, f_e, f_b, tot_op, t_i, t_e, closing_bal)
    st.download_button(
        label="📥 Download Complete Financial Statement (Landscape PDF)",
        data=stat_pdf,
        file_name=f"Landscape_Financial_Statement_{d_from}_to_{d_to}.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True
    )
    st.divider()

    t1, t2, t3 = st.tabs(["💰 Income Breakdown", "💸 Expense Breakdown", "📋 Due Collections Status"])
    with t1:
        st.dataframe(f_i, use_container_width=True)
    with t2:
        st.dataframe(f_e, use_container_width=True)
    with t3:
        st.dataframe(f_b, use_container_width=True)

# ----------------- 6. OPENING BALANCE -----------------
elif menu == "🏦 Opening Balance":
    st.subheader("🏦 Opening Balance Setup (SQL)")
    curr_c, curr_b = SQLManager.get_opening_balance()
    c1, c2 = st.columns(2)
    in_c = c1.number_input("Cash in Hand (₹)", value=float(curr_c), step=500.0)
    in_b = c2.number_input("Bank Balance (₹)", value=float(curr_b), step=500.0)
    pin = st.text_input("Enter Security PIN to Save:", type="password")
    if st.button("💾 Save Opening Balance", type="primary", use_container_width=True):
        if pin == SQLManager.get_pin():
            SQLManager.set_opening_balance(in_c, in_b)
            st.success("Opening Balance Saved to SQL!")
            st.rerun()
        else:
            st.error("Invalid PIN!")

# ----------------- 7. CUSTOMERS DIRECTORY -----------------
elif menu == "👥 Customers Directory":
    st.subheader("👥 Client Directory & Broadcast (SQL)")
    tab_new, tab_list, tab_promo = st.tabs(["➕ Add Client", "📋 Registered Clients (Edit/Delete)", "📢 Marketing / Broadcast List"])
    
    with tab_new:
        with st.form("c_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            cn = c1.text_input("Customer Name *")
            cp = c2.text_input("Mobile Number (10 Digits) *")
            
            c3, c4 = st.columns(2)
            c_addr = c3.text_input("Address / City / Village", "Kadi")
            cs = c4.selectbox("Primary Service / Purpose *", SERVICE_OPTIONS)
            
            c_notes = st.text_area("Notes / Remarks", placeholder="e.g. Visa inquiry, Land deal, Reference, etc.")
            
            if st.form_submit_button("💾 Save Client Profile", use_container_width=True):
                if cn and cp:
                    df_c = SQLManager.get_df("Customers")
                    clean_phone = str(cp).strip()
                    if not df_c.empty and "Mobile Number" in df_c and clean_phone in df_c["Mobile Number"].astype(str).values:
                        st.warning(f"⚠️ A customer with mobile {clean_phone} already exists in records!")
                    else:
                        conn = SQLManager.get_connection()
                        conn.cursor().execute("""
                            INSERT INTO Customers (Created_Date, Customer_Name, Mobile_Number, City_Address, Primary_Service, Notes)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (datetime.now().strftime("%Y-%m-%d"), cn.strip(), clean_phone, c_addr, cs, c_notes))
                        conn.commit()
                        conn.close()
                        st.success(f"Client '{cn}' saved to SQL successfully!")
                        st.rerun()
                else:
                    st.error("Customer name and mobile number are required.")
                    
    with tab_list:
        df_c = SQLManager.get_df("Customers")
        if not df_c.empty:
            search_query = st.text_input("🔍 Quick Search by Name, Mobile, Address or Service:", "")
            if search_query:
                cond = pd.Series([False]*len(df_c))
                for col in ["Customer Name", "Mobile Number", "City Address", "Primary Service", "Notes"]:
                    if col in df_c.columns:
                        cond = cond | df_c[col].astype(str).str.contains(search_query, case=False, na=False)
                filtered_df = df_c[cond]
            else:
                filtered_df = df_c
                
            st.dataframe(filtered_df, use_container_width=True)
            st.divider()
            
            sel_c_options = [f"ID #{r['ID']} - {r.get('Customer Name', '')} ({r.get('Mobile Number', '')})" for _, r in df_c.iterrows()]
            sel_c_str = st.selectbox("Select Customer to Edit / Update / Delete:", sel_c_options)
            
            if sel_c_str:
                sel_c_id = int(sel_c_str.split(" ")[1].replace("#", ""))
                c_row = df_c[df_c["ID"] == sel_c_id].iloc[0]
                
                with st.expander(f"📝 Edit Client Profile #{sel_c_id} - {c_row.get('Customer Name', '')}", expanded=True):
                    ec1, ec2 = st.columns(2)
                    u_cname = ec1.text_input("Customer Name *", str(c_row.get("Customer Name", "")))
                    u_cphone = ec2.text_input("Mobile Number *", str(c_row.get("Mobile Number", "")))
                    
                    ec3, ec4 = st.columns(2)
                    u_caddr = ec3.text_input("Address / City / Village", str(c_row.get("City Address", "Kadi")) if pd.notna(c_row.get("City Address")) else "")
                    
                    curr_serv = str(c_row.get("Primary Service", "VISA"))
                    serv_idx = SERVICE_OPTIONS.index(curr_serv) if curr_serv in SERVICE_OPTIONS else 0
                    u_cserv = ec4.selectbox("Primary Service", SERVICE_OPTIONS, index=serv_idx)
                    
                    u_cnotes = st.text_area("Notes / Remarks", str(c_row.get("Notes", "")) if pd.notna(c_row.get("Notes")) else "")
                    
                    st.markdown("🔒 **Security Confirmation:**")
                    edit_pin = st.text_input("Enter Master Security PIN:", type="password", key=f"c_pin_{sel_c_id}")
                    
                    b_col1, b_col2 = st.columns(2)
                    if b_col1.button("🔄 Update Customer Details", key=f"btn_up_{sel_c_id}", use_container_width=True):
                        if edit_pin == SQLManager.get_pin():
                            conn = SQLManager.get_connection()
                            conn.cursor().execute("""
                                UPDATE Customers SET 
                                Customer_Name = ?, Mobile_Number = ?, City_Address = ?, Primary_Service = ?, Notes = ?
                                WHERE ID = ?
                            """, (u_cname, str(u_cphone).strip(), u_caddr, u_cserv, u_cnotes, sel_c_id))
                            conn.commit()
                            conn.close()
                            st.success("Client profile updated successfully in SQL!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
                            
                    if b_col2.button("🗑️ Delete Customer", key=f"btn_del_{sel_c_id}", type="primary", use_container_width=True):
                        if edit_pin == SQLManager.get_pin():
                            conn = SQLManager.get_connection()
                            conn.cursor().execute("DELETE FROM Customers WHERE ID = ?", (sel_c_id,))
                            conn.commit()
                            conn.close()
                            st.warning("Client deleted from SQL Database!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
        else:
            st.info("No registered clients found.")

    with tab_promo:
        st.markdown("##### 📢 Bulk Broadcast & Promotion List")
        df_c = SQLManager.get_df("Customers")
        if not df_c.empty:
            serv_col_name = "Primary Service" if "Primary Service" in df_c.columns else df_c.columns[min(5, len(df_c.columns)-1)]
            service_unique_list = list(df_c[serv_col_name].dropna().unique()) if serv_col_name in df_c.columns else []
            
            sel_aud = st.selectbox("Select Target Audience:", ["All Clients"] + service_unique_list)
            target_df = df_c if (sel_aud == "All Clients" or serv_col_name not in df_c.columns) else df_c[df_c[serv_col_name] == sel_aud]
            
            st.write(f"**Total Recipients:** {len(target_df)}")
            display_cols = [c for c in ["Customer Name", "Mobile Number", "City Address", "Primary Service", "Notes"] if c in target_df.columns]
            st.dataframe(target_df[display_cols], use_container_width=True)
            
            promo_msg = st.text_area("Broadcast Message Template:", value=f"Greetings from {COMPANY_NAME}! Contact us at {COMPANY_MOBILE} for special offers and updates regarding your service inquiry.")
            for _, prow in target_df.head(10).iterrows():
                p_phone = str(prow.get("Mobile Number", "")).strip()
                p_name = prow.get("Customer Name", "Client")
                p_city = prow.get("City Address", "Kadi")
                if p_phone:
                    p_url = f"https://wa.me/91{p_phone}?text={urllib.parse.quote(promo_msg)}"
                    st.markdown(f"👉 **{p_name}** ({p_phone}) - [{p_city}]: [📲 Send WhatsApp]({p_url})")
        else:
            st.info("No client records available.")

# ----------------- 8. INCOME & EXPENSE MANAGEMENT -----------------
elif menu == "💰 Income":
    st.subheader("💰 Income Ledger (SQL)")
    df_i = SQLManager.get_df("Income")
    if not df_i.empty:
        st.dataframe(df_i, use_container_width=True)

elif menu == "💸 Expenses":
    st.subheader("💸 Expense Ledger (SQL)")
    df_e = SQLManager.get_df("Expense")
    if not df_e.empty:
        st.dataframe(df_e, use_container_width=True)

# ----------------- 9. SQL BACKUP & RESTORE -----------------
elif menu == "💾 Backup & Restore (Excel / SQL)":
    st.subheader("💾 Backup & Restore Center (Excel & SQL)")
    st.info("💡 You can export backups to Excel (.xlsx) and restore data from any previous Excel backup file.")
    
    b_tab1, b_tab2, b_tab3 = st.tabs([
        "📥 Download Backups", 
        "📤 Restore from Excel File (.xlsx)",
        "💾 Restore SQL Database File (.db)"
    ])
    
    # 1. Download Backup
    with b_tab1:
        st.markdown("##### 📥 Export All Accounting Records")
        c_bk1, c_bk2 = st.columns(2)
        
        # Excel Backup (.xlsx)
        excel_backup_data = SQLManager.export_sqlite_to_excel_buffer()
        c_bk1.download_button(
            label="📊 Download Full Excel Workbook (.xlsx)",
            data=excel_backup_data,
            file_name=f"Rojmed_Excel_Backup_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
        # SQLite DB Backup (.db)
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as f:
                c_bk2.download_button(
                    label="💾 Download SQLite Database File (.db)",
                    data=f,
                    file_name=f"rojmed_sql_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.db",
                    mime="application/x-sqlite3",
                    use_container_width=True
                )

    # 2. Restore from Excel (.xlsx)
    with b_tab2:
        st.markdown("##### 📤 Upload & Restore from Excel File (.xlsx)")
        st.caption("Upload your previously downloaded Excel backup file. All sheets will be imported back into SQL Database.")
        
        uploaded_excel = st.file_uploader("Choose Backup Excel File (.xlsx):", type=["xlsx"], key="restore_excel_uploader")
        excel_pin = st.text_input("Enter Master Security PIN to Confirm Restore:", type="password", key="rest_pin_excel")
        
        if st.button("🚀 Restore Data from Excel File", type="primary", use_container_width=True):
            if excel_pin == SQLManager.get_pin():
                if uploaded_excel is not None:
                    success, msg = SQLManager.import_excel_to_sqlite(uploaded_excel)
                    if success:
                        st.success("✅ Excel data successfully imported and restored into SQL Database!")
                        st.rerun()
                    else:
                        st.error(f"❌ Error restoring Excel data: {msg}")
                else:
                    st.error("Please choose a valid `.xlsx` Excel backup file.")
            else:
                st.error("❌ Incorrect Security PIN! Action denied.")

    # 3. Restore from SQLite (.db)
    with b_tab3:
        st.markdown("##### 💾 Restore SQLite Database File (.db)")
        uploaded_db = st.file_uploader("Choose SQLite Database File (.db):", type=["db", "sqlite", "sqlite3"], key="restore_db_uploader")
        db_pin = st.text_input("Enter Master Security PIN:", type="password", key="rest_pin_db")
        
        if st.button("🚀 Restore SQL Database", type="primary", use_container_width=True):
            if db_pin == SQLManager.get_pin():
                if uploaded_db is not None:
                    with open(DB_FILE, "wb") as f:
                        f.write(uploaded_db.getbuffer())
                    st.success("✅ SQL Database file restored!")
                    st.rerun()
                else:
                    st.error("Please choose a valid `.db` file.")
            else:
                st.error("❌ Incorrect Security PIN!")

# ----------------- 10. SECURITY PIN -----------------
elif menu == "⚙️ Security / Change PIN":
    st.subheader("⚙️ Change Master PIN")
    with st.form("pin_form"):
        old_p = st.text_input("Current PIN *", type="password")
        new_p = st.text_input("New PIN *", type="password")
        conf_p = st.text_input("Confirm New PIN *", type="password")
        if st.form_submit_button("💾 Update PIN"):
            if old_p == SQLManager.get_pin():
                if new_p and new_p == conf_p:
                    SQLManager.set_pin(new_p)
                    st.success("PIN Updated in SQL Settings!")
                else:
                    st.error("PIN mismatch!")
            else:
                st.error("Incorrect Current PIN!")
