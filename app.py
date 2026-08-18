import os
import urllib.parse
from datetime import datetime, time
import io
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl import Workbook

# ReportLab for PDF Bill & Reports
from reportlab.lib.pagesizes import A4
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

# Security PIN File
PIN_FILE = "security_pin.txt"
DEFAULT_PIN = "1234"

def get_saved_pin():
    if not os.path.exists(PIN_FILE):
        with open(PIN_FILE, "w") as f:
            f.write(DEFAULT_PIN)
        return DEFAULT_PIN
    with open(PIN_FILE, "r") as f:
        return f.read().strip()

def save_new_pin(new_pin):
    with open(PIN_FILE, "w") as f:
        f.write(str(new_pin).strip())

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
            transition: all 0.2s ease-in-out;
            border: 1px solid #CBD5E1;
            background-color: #FFFFFF;
            color: #1E293B;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }
        .stButton>button:hover {
            border-color: #2563EB;
            color: #2563EB;
            background-color: #F1F5F9;
            transform: translateY(-1px);
        }
        
        button[kind="primary"] {
            background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
            color: #FFFFFF !important;
            border: none !important;
            box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25) !important;
        }
        button[kind="primary"]:hover {
            background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%) !important;
            color: #FFFFFF !important;
            box-shadow: 0 4px 10px rgba(37, 99, 235, 0.35) !important;
        }

        div[data-testid="metric-container"] {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            padding: 14px 18px;
            border-radius: 10px;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.03);
            border-left: 4px solid #2563EB;
        }
        div[data-testid="stMetricValue"] {
            font-weight: 700;
            color: #0F172A;
        }
        
        .stDataFrame, div[data-testid="stExpander"] {
            background-color: #FFFFFF;
            border-radius: 8px;
            border: 1px solid #E2E8F0;
        }
        
        .stTextInput>div>div>input, .stNumberInput>div>div>input, .stSelectbox>div>div {
            border-radius: 8px;
            border: 1px solid #CBD5E1;
            background-color: #FFFFFF;
            color: #0F172A;
        }
    </style>
""", unsafe_allow_html=True)

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
                    <h2 style="color: #1E3A8A; margin: 0; font-size: 24px; font-weight: 800; letter-spacing: 0.5px;">{COMPANY_NAME}</h2>
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

# ----------------- LOGIN AUTHENTICATION CHECK -----------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    render_top_logos()
    
    col_c1, col_c2, col_c3 = st.columns([1, 1.2, 1])
    with col_c2:
        st.markdown("""
            <div style="background: #FFFFFF; padding: 25px; border-radius: 12px; border: 1px solid #E2E8F0; box-shadow: 0 4px 15px rgba(0,0,0,0.05); text-align: center;">
                <h3 style="color: #1E3A8A; margin-top: 0;">🔒 Secure System Login</h3>
                <p style="color: #64748B; font-size: 13px;">Enter Master Security PIN to unlock the ledger.</p>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        entered_pin = st.text_input("Enter 4-Digit Security PIN:", type="password", max_chars=8)
        
        if st.button("🔓 Unlock & Login", type="primary", use_container_width=True):
            if entered_pin == get_saved_pin():
                st.session_state.logged_in = True
                st.success("Access Granted! Loading ledger...")
                st.rerun()
            else:
                st.error("❌ Invalid Security PIN! Please try again.")
        st.caption(f"Default setup PIN is: `{DEFAULT_PIN}` (You can change it inside settings).")
    st.stop()

# Excel Data Manager
EXCEL_FILE = "Rojmed_Data.xlsx"

class ExcelManager:
    @staticmethod
    def init_excel():
        if not os.path.exists(EXCEL_FILE):
            wb = Workbook()
            wb.remove(wb.active)
            sheets = {
                "Settings": ["Key", "Value", "Updated Date"],
                "Invoices_Archive": ["Invoice No", "Date", "Customer Name", "Mobile Number", "Service 1", "Amount 1", "Service 2", "Amount 2", "Total Amount", "Paid Amount", "Pending Amount", "Payment Mode", "Remarks"],
                "Customers": ["ID", "Created Date", "Customer Name", "Mobile Number", "City/Address", "Primary Service / Purpose", "Notes"],
                "Income": ["ID", "Date", "Customer/Person", "Work Details", "Amount", "Payment Mode", "Notes"],
                "Expense": ["ID", "Date", "Expense Name", "Amount", "Notes"],
                "Udhar_Baki": ["ID", "Date", "Customer Name", "Mobile Number", "Service Details", "Total Amount", "Paid Amount", "Pending Amount", "Due Date", "Status"],
                "Task_Reminder": ["ID", "Date", "Time", "Person Name", "Mobile", "Task Details", "Status"]
            }
            for sheet_name, headers in sheets.items():
                ws = wb.create_sheet(title=sheet_name)
                ws.append(headers)
            
            ws_set = wb["Settings"]
            ws_set.append(["Cash_Opening_Balance", 0.0, datetime.now().strftime("%Y-%m-%d")])
            ws_set.append(["Bank_Opening_Balance", 0.0, datetime.now().strftime("%Y-%m-%d")])
            
            wb.save(EXCEL_FILE)

    @staticmethod
    def get_df(sheet_name):
        ExcelManager.init_excel()
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
            return df
        except Exception:
            return pd.DataFrame()

    @staticmethod
    def save_df(sheet_name, df):
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    @staticmethod
    def append_row(sheet_name, row_dict):
        df = ExcelManager.get_df(sheet_name)
        new_id = 1 if df.empty or "ID" not in df else (df["ID"].max() + 1 if not pd.isna(df["ID"].max()) else len(df) + 1)
        if "ID" in df.columns or sheet_name not in ["Invoices_Archive", "Settings"]:
            row_dict["ID"] = int(new_id)
        new_row_df = pd.DataFrame([row_dict])
        df = pd.concat([df, new_row_df], ignore_index=True)
        ExcelManager.save_df(sheet_name, df)
        return int(new_id)

    @staticmethod
    def delete_row(sheet_name, row_id):
        df = ExcelManager.get_df(sheet_name)
        if not df.empty and "ID" in df:
            df = df[df["ID"] != row_id]
            ExcelManager.save_df(sheet_name, df)

    @staticmethod
    def update_row(sheet_name, row_id, updated_dict):
        df = ExcelManager.get_df(sheet_name)
        if not df.empty and "ID" in df:
            idx = df.index[df["ID"] == row_id].tolist()
            if idx:
                for k, v in updated_dict.items():
                    df.at[idx[0], k] = v
                ExcelManager.save_df(sheet_name, df)

    @staticmethod
    def get_opening_balance():
        df_set = ExcelManager.get_df("Settings")
        cash_op = 0.0
        bank_op = 0.0
        if not df_set.empty and "Key" in df_set and "Value" in df_set:
            cash_row = df_set[df_set["Key"] == "Cash_Opening_Balance"]
            bank_row = df_set[df_set["Key"] == "Bank_Opening_Balance"]
            if not cash_row.empty:
                cash_op = float(cash_row["Value"].values[0]) if pd.notna(cash_row["Value"].values[0]) else 0.0
            if not bank_row.empty:
                bank_op = float(bank_row["Value"].values[0]) if pd.notna(bank_row["Value"].values[0]) else 0.0
        return cash_op, bank_op

    @staticmethod
    def set_opening_balance(cash_op, bank_op):
        now_str = datetime.now().strftime("%Y-%m-%d")
        new_data = [
            {"Key": "Cash_Opening_Balance", "Value": float(cash_op), "Updated Date": now_str},
            {"Key": "Bank_Opening_Balance", "Value": float(bank_op), "Updated Date": now_str}
        ]
        df_new = pd.DataFrame(new_data)
        ExcelManager.save_df("Settings", df_new)

# Initialize Database & Render Header
ExcelManager.init_excel()
render_top_logos()

# ----------------- BUTTON-STYLE SIDEBAR NAVIGATION -----------------
if "current_page" not in st.session_state:
    st.session_state.current_page = "📊 Dashboard"

st.sidebar.markdown("<h4 style='color:#1E3A8A; margin-bottom: 0px;'>📌 Navigation Menu</h4>", unsafe_allow_html=True)
st.sidebar.markdown("<p style='font-size:12px; color:#64748B; margin-top:2px;'>Select active screen below</p>", unsafe_allow_html=True)

menu_items = [
    ("📊 Dashboard", "Business metrics & live summary"),
    ("⏰ Task Reminders", "Calendar & Clock Reminders & Task Tracking"),
    ("🧾 Generate Bill / Voucher", "Create, print, re-print invoices & post entries"),
    ("📄 Reports & PDF", "Statements for Income, Expense & Pending Dues"),
    ("🏦 Opening Balance", "Set starting cash, bank & old customer dues"),
    ("👥 Customers Directory", "Add/Manage clients & marketing list"),
    ("💰 Income", "View, edit, and manage income records"),
    ("💸 Expenses", "View, edit, and manage expense records"),
    ("📋 Due Collections", "Customer pending dues & settlements"),
    ("⚙️ Security / Change PIN", "Change Master Login & Edit PIN")
]

for label, desc in menu_items:
    is_active = st.session_state.current_page == label
    btn_type = "primary" if is_active else "secondary"
    if st.sidebar.button(label, key=f"nav_{label}", type=btn_type, use_container_width=True):
        st.session_state.current_page = label
        st.rerun()

menu = st.session_state.current_page

st.sidebar.markdown("---")
if st.sidebar.button("🔒 Logout System", use_container_width=True):
    st.session_state.logged_in = False
    st.rerun()

st.sidebar.info(f"📍 **Location:** Kadi, Gujarat\n📞 **Helpline:** {COMPANY_MOBILE}")

# Master Service List
SERVICE_OPTIONS = [
    "VISA",
    "PASSPORT",
    "PANCARD",
    "LOAN",
    "LAND RECORD (7/12 & 8-A)",
    "E-KYC ALL",
    "FARMER REGISTRATION",
    "PM-KISAN",
    "DIGITAL GUJARAT SERVICES",
    "PROVIDENT FUND (PF)",
    "REVENUE WORK",
    "MARRIAGE CERTIFICATE",
    "INSURANCE",
    "AIR TICKET",
    "OTHER"
]

def get_customer_pending_due(phone_number):
    if not phone_number:
        return 0.0, []
    df_baki = ExcelManager.get_df("Udhar_Baki")
    if df_baki.empty:
        return 0.0, []
    mob_col = "Mobile Number" if "Mobile Number" in df_baki else "Mobile"
    baki_col = "Pending Amount" if "Pending Amount" in df_baki else "Amount"
    clean_p = str(phone_number).strip()
    cust_dues = df_baki[(df_baki[mob_col].astype(str).str.strip() == clean_p) & (df_baki[baki_col] > 0)]
    total_due = cust_dues[baki_col].sum() if not cust_dues.empty else 0.0
    return float(total_due), cust_dues

def generate_invoice_pdf_buffer(bill_no, bill_date, cust_name, cust_phone, service_1, amt_1, service_2, amt_2, total_bill, rec_amt, baki_amt, pay_mode, remarks):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=20, bottomMargin=20)
    elems = []
    styles = getSampleStyleSheet()
    
    c_title = ParagraphStyle('C1', parent=styles['Heading1'], fontSize=15, leading=17, textColor=colors.HexColor("#1E3A8A"), alignment=1)
    c_sub = ParagraphStyle('C2', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor("#334155"), alignment=1)
    c_inv = ParagraphStyle('C3', parent=styles['Heading2'], fontSize=12, leading=14, textColor=colors.HexColor("#15803D"), alignment=1)
    
    logo_l = None
    logo_r = None
    if os.path.exists(LOGO_VISA):
        logo_l = PDFImage(LOGO_VISA, width=55, height=55)
    if os.path.exists(LOGO_FINCARE):
        logo_r = PDFImage(LOGO_FINCARE, width=80, height=40)
    
    header_center_text = [
        Paragraph(f"<b>{COMPANY_NAME}</b>", c_title),
        Paragraph(f"📍 {COMPANY_ADDRESS}", c_sub),
        Paragraph(f"<b>📞 Phone:</b> {COMPANY_MOBILE}", c_sub),
        Paragraph(f"<i>{COMPANY_TAGLINE}</i>", c_sub)
    ]
    
    hdr_table_data = [[
        logo_l if logo_l else "",
        header_center_text,
        logo_r if logo_r else ""
    ]]
    t_hdr = Table(hdr_table_data, colWidths=[70, 400, 80])
    t_hdr.setStyle(TableStyle([
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
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
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
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
    t_items.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ('ALIGN', (-1, 0), (-1, -1), 'RIGHT'),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#F0FDF4")),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    elems.append(t_items)
    elems.append(Spacer(1, 15))
    
    foot_data = [
        [f"<b>Note:</b> {remarks}", f"For, <b>{COMPANY_NAME}</b>"],
        ["", "\n\n___________________________\nAuthorized Signature"]
    ]
    t_foot = Table([[Paragraph(c, styles['Normal']) for c in r] for r in foot_data], colWidths=[310, 240])
    t_foot.setStyle(TableStyle([
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    elems.append(t_foot)
    
    doc.build(elems)
    buf.seek(0)
    return buf

# ----------------- 1. DASHBOARD WITH LIVE NOTIFICATIONS & REMINDERS -----------------
if menu == "📊 Dashboard":
    st.subheader("📊 Business Overview & Live Summary")
    
    df_rem_all = ExcelManager.get_df("Task_Reminder")
    if not df_rem_all.empty and "Status" in df_rem_all:
        pending_tasks = df_rem_all[df_rem_all["Status"] == "Pending"]
        if not pending_tasks.empty:
            st.markdown(f"""
                <div style="background: #FEF2F2; border-left: 5px solid #EF4444; padding: 12px 18px; border-radius: 8px; margin-bottom: 15px;">
                    <h4 style="color: #B91C1C; margin: 0 0 6px 0;">🔔 Active Reminders & Pending Tasks ({len(pending_tasks)})</h4>
                    <p style="color: #7F1D1D; font-size: 13px; margin: 0;">Below are upcoming scheduled follow-ups and pending client works.</p>
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
                    ExcelManager.update_row("Task_Reminder", t_id, {"Status": "Completed"})
                    st.success(f"Task #{t_id} marked as Completed!")
                    st.rerun()
            st.divider()

    df_inc = ExcelManager.get_df("Income")
    df_exp = ExcelManager.get_df("Expense")
    df_baki = ExcelManager.get_df("Udhar_Baki")
    df_cust = ExcelManager.get_df("Customers")
    
    today_str = datetime.now().strftime("%Y-%m-%d")
    
    inc_date_col = "Date" if "Date" in df_inc else None
    inc_amt_col = "Amount" if "Amount" in df_inc else None
    exp_date_col = "Date" if "Date" in df_exp else None
    exp_amt_col = "Amount" if "Amount" in df_exp else None
    baki_amt_col = "Pending Amount" if "Pending Amount" in df_baki else None
    
    today_inc = df_inc[df_inc[inc_date_col] == today_str][inc_amt_col].sum() if not df_inc.empty and inc_date_col and inc_amt_col else 0
    total_inc = df_inc[inc_amt_col].sum() if not df_inc.empty and inc_amt_col else 0
    today_exp = df_exp[df_exp[exp_date_col] == today_str][exp_amt_col].sum() if not df_exp.empty and exp_date_col and exp_amt_col else 0
    total_exp = df_exp[exp_amt_col].sum() if not df_exp.empty and exp_amt_col else 0
    total_baki = df_baki[baki_amt_col].sum() if not df_baki.empty and baki_amt_col else 0
    
    cash_op, bank_op = ExcelManager.get_opening_balance()
    total_opening_balance = cash_op + bank_op
    closing_net_balance = total_opening_balance + total_inc - total_exp
    total_cust = len(df_cust) if not df_cust.empty else 0

    col0, col1, col2, col3, col4, col5 = st.columns(6)
    col0.metric("Opening Balance", f"₹ {total_opening_balance:,.2f}", f"Cash: ₹{cash_op:,.0f} | Bank: ₹{bank_op:,.0f}")
    col1.metric("Today's Income", f"₹ {today_inc:,.2f}", f"Total: ₹ {total_inc:,.2f}")
    col2.metric("Today's Expense", f"₹ {today_exp:,.2f}", f"Total: ₹ {total_exp:,.2f}", delta_color="inverse")
    col3.metric("Closing Balance", f"₹ {closing_net_balance:,.2f}")
    col4.metric("Total Pending Dues", f"₹ {total_baki:,.2f}")
    col5.metric("Registered Clients", f"{total_cust}")

    st.divider()
    st.subheader("📋 Pending Collections & Itemized WhatsApp Reminders")
    
    if not df_baki.empty and baki_amt_col:
        name_col = "Customer Name" if "Customer Name" in df_baki else "Name"
        mob_col = "Mobile Number" if "Mobile Number" in df_baki else "Mobile"
        due_col = "Due Date" if "Due Date" in df_baki else "Date"
        serv_col = "Service Details" if "Service Details" in df_baki else ("Primary Service / Purpose" if "Primary Service / Purpose" in df_baki else None)
        
        pending = df_baki[df_baki[baki_amt_col] > 0]
        if not pending.empty:
            for _, r in pending.iterrows():
                b1, b2, b3, b4, b5, b6 = st.columns([2, 2, 2, 2, 2, 2])
                b1.write(f"**{r[name_col]}**")
                b2.write(f"📞 {r[mob_col]}")
                
                serv_name = str(r[serv_col]) if serv_col and pd.notna(r.get(serv_col)) else "Service / Work"
                b3.write(f"🏷️ *{serv_name}*")
                b4.write(f"Due: **₹ {r[baki_amt_col]:,.2f}**")
                b5.write(f"Date: {r[due_col]}")
                
                msg = f"Hello {r[name_col]}, gentle payment reminder from {COMPANY_NAME}. An outstanding balance of Rs. {r[baki_amt_col]:,.2f} is pending for {serv_name}. Kindly settle at your earliest convenience. Contact: {COMPANY_MOBILE}"
                wa_url = f"https://wa.me/91{str(r[mob_col]).strip()}?text={urllib.parse.quote(msg)}"
                b6.markdown(f"[📲 Send WhatsApp]({wa_url})", unsafe_allow_html=True)
        else:
            st.success("All customer accounts are clear! No pending dues.")
    else:
        st.info("No credit or pending records found.")

# ----------------- 2. TASK REMINDERS (CALENDAR & CLOCK + COMPLETE BUTTON) -----------------
elif menu == "⏰ Task Reminders":
    st.subheader("⏰ Reminders & Task Management (Calendar & Clock)")
    
    tab_new_task, tab_pending_tasks, tab_completed_tasks = st.tabs(["➕ Schedule New Reminder", "⏳ Active / Pending Tasks", "✅ Completed Task History"])
    
    with tab_new_task:
        with st.form("task_picker_form", clear_on_submit=True):
            col_t1, col_t2 = st.columns(2)
            task_desc = col_t1.text_input("Task / Follow-up Details *", placeholder="e.g. Appointment for Visa File, Document Collection")
            person_name = col_t2.text_input("Associated Person / Customer Name *")
            
            col_t3, col_t4 = st.columns(2)
            rem_phone = col_t3.text_input("Mobile Number (10 Digits)", placeholder="e.g. 9876543210")
            
            col_cal, col_clk = st.columns(2)
            rem_date = col_cal.date_input("📅 Select Reminder Date *", datetime.now()).strftime("%Y-%m-%d")
            rem_time = col_clk.time_input("⏰ Select Reminder Time *", time(11, 0)).strftime("%I:%M %p")
            
            if st.form_submit_button("💾 Save & Schedule Reminder", use_container_width=True):
                if task_desc and person_name:
                    ExcelManager.append_row("Task_Reminder", {
                        "Date": rem_date,
                        "Time": rem_time,
                        "Person Name": person_name,
                        "Mobile": rem_phone,
                        "Task Details": task_desc,
                        "Status": "Pending"
                    })
                    st.success(f"✅ Reminder scheduled for {rem_date} at {rem_time}!")
                    st.rerun()
                else:
                    st.error("Task details and person name are mandatory.")

    with tab_pending_tasks:
        df_rem = ExcelManager.get_df("Task_Reminder")
        if not df_rem.empty:
            pending_list = df_rem[df_rem["Status"] == "Pending"]
            if not pending_list.empty:
                st.write(f"**Total Pending Reminders:** {len(pending_list)}")
                for _, r in pending_list.iterrows():
                    r_id = r["ID"]
                    with st.container():
                        c1, c2, c3, c4, c5 = st.columns([2, 1.5, 3, 2, 2])
                        c1.write(f"📅 **{r['Date']}**")
                        c2.write(f"⏰ {r['Time']}")
                        c3.write(f"📌 **{r['Task Details']}**\n👤 {r.get('Person Name', '-')}")
                        
                        if pd.notna(r.get('Mobile')):
                            w_msg = f"Reminder regarding {r['Task Details']} scheduled on {r['Date']} at {r['Time']}."
                            w_url = f"https://wa.me/91{str(r.get('Mobile')).strip()}?text={urllib.parse.quote(w_msg)}"
                            c4.markdown(f"[📲 WhatsApp]({w_url})")
                        else:
                            c4.write("-")
                            
                        if c5.button("✅ Mark Completed", key=f"rem_done_{r_id}"):
                            ExcelManager.update_row("Task_Reminder", r_id, {"Status": "Completed"})
                            st.success("Task marked as Completed!")
                            st.rerun()
                        st.markdown("<hr style='margin:6px 0; border:0; height:1px; background:#E2E8F0;'>", unsafe_allow_html=True)
            else:
                st.success("✨ No pending reminders! All tasks are up to date.")
        else:
            st.info("No task reminders found.")

    with tab_completed_tasks:
        df_rem = ExcelManager.get_df("Task_Reminder")
        if not df_rem.empty:
            done_list = df_rem[df_rem["Status"] == "Completed"]
            if not done_list.empty:
                st.dataframe(done_list, use_container_width=True)
            else:
                st.info("No completed tasks yet.")
        else:
            st.info("No tasks recorded.")

# ----------------- 3. BILL / INVOICE & RE-PRINT + WHATSAPP SHARE -----------------
elif menu == "🧾 Generate Bill / Voucher":
    st.subheader("🧾 Generate Invoice, Voucher & Re-Print Old Bills")
    
    bill_type = st.radio("Select Action:", ["Customer Invoice (Income)", "🖨️ Re-Print Old Invoice / Voucher", "Payment Voucher (Expense)", "Settle Old Pending Due (Direct Payment)"], horizontal=True)
    st.divider()
    
    # ----------------- MODE A: CUSTOMER INVOICE -----------------
    if bill_type == "Customer Invoice (Income)":
        st.markdown("### 👤 STEP 1: Client Selection & Verification")
        df_cust = ExcelManager.get_df("Customers")
        
        search_mode = st.radio("Choose Mode:", ["Select Existing Customer from Directory", "Enter New Customer Directly"], horizontal=True)
        
        pre_name = ""
        pre_phone = ""
        suggested_service = "VISA"
        
        if search_mode == "Select Existing Customer from Directory" and not df_cust.empty:
            cust_options = ["-- Select Client --"] + [f"{r['Customer Name']} ({r['Mobile Number']}) - {r.get('Primary Service / Purpose', '')}" for _, r in df_cust.iterrows()]
            selected_cust = st.selectbox("🔍 Search & Choose Registered Client:", cust_options)
            
            if selected_cust != "-- Select Client --":
                chosen_idx = cust_options.index(selected_cust) - 1
                row_c = df_cust.iloc[chosen_idx]
                pre_name = str(row_c["Customer Name"])
                pre_phone = str(row_c["Mobile Number"])
                cust_saved_service = str(row_c.get("Primary Service / Purpose", "VISA"))
                suggested_service = cust_saved_service if cust_saved_service in SERVICE_OPTIONS else "OTHER"
                st.success(f"✅ Verified Customer: **{pre_name}** | Primary Purpose: *{cust_saved_service}*")
        
        col_c1, col_c2 = st.columns(2)
        cust_name = col_c1.text_input("Customer Name *", value=pre_name)
        cust_phone = col_c2.text_input("Mobile Number *", value=pre_phone)
        
        if cust_phone:
            old_due, old_due_df = get_customer_pending_due(cust_phone)
            if old_due > 0:
                st.error(f"⚠️ **Old Due Alert:** This customer has previous pending balance of **₹ {old_due:,.2f}**!")
                with st.expander("View Customer's Pending Due Records"):
                    st.dataframe(old_due_df[["Date", "Service Details", "Total Amount", "Paid Amount", "Pending Amount", "Due Date"]], use_container_width=True)
            else:
                st.info("✨ Customer Account Status: All previous dues clear.")
        
        st.markdown("---")
        
        st.markdown("### 🧾 STEP 2: Bill & Service Particulars")
        col_in1, col_in2 = st.columns(2)
        bill_no = col_in1.text_input("Invoice No.", f"INV-{datetime.now().strftime('%Y%m%d%H%M')}")
        bill_date = col_in2.date_input("Invoice Date", datetime.now()).strftime("%Y-%m-%d")
        
        st.markdown("##### 💼 Service 1 (Main Service)")
        col_s1, col_s2 = st.columns(2)
        default_idx = SERVICE_OPTIONS.index(suggested_service) if suggested_service in SERVICE_OPTIONS else 0
        service_1_sel = col_s1.selectbox("Select Service 1 *", SERVICE_OPTIONS, index=default_idx, key="s1_sel")
        service_1 = col_s1.text_input("Type Custom Service 1 Name *", key="s1_custom") if service_1_sel == "OTHER" else service_1_sel
        amt_1 = col_s2.number_input("Amount 1 (₹) *", min_value=0.0, step=100.0, key="amt1")
        
        st.markdown("##### 💼 Service 2 (Optional Additional Service)")
        col_s3, col_s4 = st.columns(2)
        service_2_sel = col_s3.selectbox("Select Service 2 (Optional)", ["None"] + SERVICE_OPTIONS, index=0, key="s2_sel")
        if service_2_sel == "OTHER":
            service_2 = col_s3.text_input("Type Custom Service 2 Name", key="s2_custom")
        elif service_2_sel == "None":
            service_2 = ""
        else:
            service_2 = service_2_sel
            
        amt_2 = col_s4.number_input("Amount 2 (₹)", min_value=0.0, step=100.0, key="amt2")
        
        total_bill = amt_1 + amt_2
        st.write(f"### **Total Current Bill: ₹ {total_bill:,.2f}**")
        
        col_p1, col_p2, col_p3 = st.columns(3)
        pay_mode = col_p1.selectbox("Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque", "Pending / Due"])
        rec_amt = col_p2.number_input("Received Amount (₹)", min_value=0.0, max_value=float(total_bill), value=float(total_bill) if pay_mode != "Pending / Due" else 0.0, step=100.0)
        due_date = col_p3.date_input("Due Date (If balance pending)", datetime.now()).strftime("%Y-%m-%d")
        
        baki_amt = total_bill - rec_amt
        item_due_desc = service_1 + (f" + {service_2}" if service_2 else "")
        
        if baki_amt > 0:
            st.warning(f"⚠️ Balance Due: ₹ {baki_amt:,.2f} for '{item_due_desc}' (Will be logged under Due Collections)")
        
        remarks = st.text_input("Remarks / Notes", "Thank you for choosing our services!")
        
        col_chk1, col_chk2 = st.columns(2)
        auto_save = col_chk1.checkbox("☑ Automatically post this entry to Income & Due Collections", value=True)
        auto_save_cust = col_chk2.checkbox("☑ Auto-save customer into directory if new", value=True)

        if st.button("💾 Generate Bill, Save & Export PDF", type="primary", use_container_width=True):
            if cust_name and total_bill > 0 and service_1:
                ExcelManager.append_row("Invoices_Archive", {
                    "Invoice No": bill_no,
                    "Date": bill_date,
                    "Customer Name": cust_name,
                    "Mobile Number": str(cust_phone).strip(),
                    "Service 1": service_1,
                    "Amount 1": amt_1,
                    "Service 2": service_2,
                    "Amount 2": amt_2,
                    "Total Amount": total_bill,
                    "Paid Amount": rec_amt,
                    "Pending Amount": baki_amt,
                    "Payment Mode": pay_mode,
                    "Remarks": remarks
                })
                
                if auto_save_cust:
                    df_c = ExcelManager.get_df("Customers")
                    clean_phone = str(cust_phone).strip()
                    if df_c.empty or "Mobile Number" not in df_c or clean_phone not in df_c["Mobile Number"].astype(str).values:
                        ExcelManager.append_row("Customers", {
                            "Created Date": bill_date,
                            "Customer Name": cust_name,
                            "Mobile Number": clean_phone,
                            "City/Address": "Kadi",
                            "Primary Service / Purpose": service_1,
                            "Notes": remarks
                        })
                
                if auto_save:
                    if rec_amt > 0:
                        work_desc = f"{service_1}" + (f", {service_2}" if service_2 else "")
                        ExcelManager.append_row("Income", {
                            "Date": bill_date,
                            "Customer/Person": cust_name,
                            "Work Details": f"Bill #{bill_no}: {work_desc}",
                            "Amount": rec_amt,
                            "Payment Mode": pay_mode,
                            "Notes": f"Mob: {cust_phone} | {remarks}"
                        })
                    if baki_amt > 0:
                        ExcelManager.append_row("Udhar_Baki", {
                            "Date": bill_date,
                            "Customer Name": cust_name,
                            "Mobile Number": cust_phone,
                            "Service Details": item_due_desc,
                            "Total Amount": total_bill,
                            "Paid Amount": rec_amt,
                            "Pending Amount": baki_amt,
                            "Due Date": due_date,
                            "Status": "Pending"
                        })
                    st.success("✅ Entry posted to Income & Due Collections automatically!")

                pdf_data = generate_invoice_pdf_buffer(bill_no, bill_date, cust_name, cust_phone, service_1, amt_1, service_2, amt_2, total_bill, rec_amt, baki_amt, pay_mode, remarks)
                
                col_dwn, col_wa = st.columns(2)
                col_dwn.download_button(
                    label=f"📥 Download PDF Invoice ({bill_no})",
                    data=pdf_data,
                    file_name=f"Invoice_{cust_name}_{bill_no}.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
                
                wa_bill_text = (
                    f"🧾 *TAX INVOICE / RECEIPT*\n"
                    f"🏢 *{COMPANY_NAME}*\n"
                    f"📍 {COMPANY_ADDRESS}\n"
                    f"📞 {COMPANY_MOBILE}\n"
                    f"----------------------------\n"
                    f"📄 *Invoice No:* {bill_no}\n"
                    f"📅 *Date:* {bill_date}\n"
                    f"👤 *Customer:* {cust_name}\n"
                    f"💼 *Service:* {item_due_desc}\n"
                    f"💰 *Total Amount:* Rs. {total_bill:,.2f}\n"
                    f"✅ *Paid Amount:* Rs. {rec_amt:,.2f} ({pay_mode})\n"
                )
                if baki_amt > 0:
                    wa_bill_text += f"⚠️ *Pending Due:* Rs. {baki_amt:,.2f} (Due: {due_date})\n"
                wa_bill_text += f"----------------------------\n🙏 *Thank you for your business!*"
                
                clean_target_phone = str(cust_phone).strip()
                wa_link_url = f"https://wa.me/91{clean_target_phone}?text={urllib.parse.quote(wa_bill_text)}"
                col_wa.markdown(f'<a href="{wa_link_url}" target="_blank"><button style="width:100%; height:45px; background-color:#25D366; color:white; border:none; border-radius:8px; font-weight:bold; font-size:15px; cursor:pointer;">📲 Send Invoice via WhatsApp</button></a>', unsafe_allow_html=True)
            else:
                st.error("Please enter customer name, valid service, and bill amount.")

    # ----------------- MODE B: RE-PRINT OLD INVOICES -----------------
    elif bill_type == "🖨️ Re-Print Old Invoice / Voucher":
        st.markdown("### 🖨️ Re-Print & WhatsApp Share Previous Invoices")
        df_arch = ExcelManager.get_df("Invoices_Archive")
        
        if not df_arch.empty and "Invoice No" in df_arch:
            search_inv = st.text_input("🔍 Search Archive by Invoice No, Customer Name, or Phone:", "")
            if search_inv:
                f_arch = df_arch[
                    df_arch["Invoice No"].astype(str).str.contains(search_inv, case=False, na=False) |
                    df_arch["Customer Name"].astype(str).str.contains(search_inv, case=False, na=False) |
                    df_arch["Mobile Number"].astype(str).str.contains(search_inv, case=False, na=False)
                ]
            else:
                f_arch = df_arch
                
            st.dataframe(f_arch, use_container_width=True)
            st.divider()
            
            sel_inv_options = [f"{r['Invoice No']} - {r['Customer Name']} ({r['Date']}) | ₹{r['Total Amount']}" for _, r in f_arch.iterrows()]
            chosen_inv_str = st.selectbox("Select Invoice to Re-Print / WhatsApp:", sel_inv_options)
            
            if chosen_inv_str:
                sel_inv_no = chosen_inv_str.split(" - ")[0]
                inv_row = df_arch[df_arch["Invoice No"] == sel_inv_no].iloc[0]
                
                re_pdf = generate_invoice_pdf_buffer(
                    str(inv_row["Invoice No"]),
                    str(inv_row["Date"]),
                    str(inv_row["Customer Name"]),
                    str(inv_row["Mobile Number"]),
                    str(inv_row.get("Service 1", "Service")),
                    float(inv_row.get("Amount 1", 0.0)),
                    str(inv_row.get("Service 2", "")) if pd.notna(inv_row.get("Service 2")) else "",
                    float(inv_row.get("Amount 2", 0.0)) if pd.notna(inv_row.get("Amount 2")) else 0.0,
                    float(inv_row["Total Amount"]),
                    float(inv_row.get("Paid Amount", inv_row["Total Amount"])),
                    float(inv_row.get("Pending Amount", 0.0)),
                    str(inv_row.get("Payment Mode", "Cash")),
                    str(inv_row.get("Remarks", "Thank you for choosing our services!"))
                )
                
                col_dwn2, col_wa2 = st.columns(2)
                col_dwn2.download_button(
                    label=f"🖨️ Re-Download Invoice PDF ({sel_inv_no})",
                    data=re_pdf,
                    file_name=f"RePrint_{sel_inv_no}.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
                
                serv_comb = str(inv_row.get('Service 1', 'Service')) + (f" + {inv_row.get('Service 2')}" if pd.notna(inv_row.get('Service 2')) and str(inv_row.get('Service 2')).strip() else "")
                wa_re_msg = (
                    f"🧾 *TAX INVOICE / RECEIPT (COPY)*\n"
                    f"🏢 *{COMPANY_NAME}*\n"
                    f"📍 {COMPANY_ADDRESS}\n"
                    f"📞 {COMPANY_MOBILE}\n"
                    f"----------------------------\n"
                    f"📄 *Invoice No:* {inv_row['Invoice No']}\n"
                    f"📅 *Date:* {inv_row['Date']}\n"
                    f"👤 *Customer:* {inv_row['Customer Name']}\n"
                    f"💼 *Service:* {serv_comb}\n"
                    f"💰 *Total Amount:* Rs. {float(inv_row['Total Amount']):,.2f}\n"
                    f"✅ *Paid Amount:* Rs. {float(inv_row.get('Paid Amount', inv_row['Total Amount'])):,.2f}\n"
                )
                if float(inv_row.get("Pending Amount", 0.0)) > 0:
                    wa_re_msg += f"⚠️ *Pending Due:* Rs. {float(inv_row.get('Pending Amount', 0.0)):,.2f}\n"
                wa_re_msg += f"----------------------------\n🙏 *Thank you for your business!*"
                
                re_target_phone = str(inv_row["Mobile Number"]).strip()
                re_wa_url = f"https://wa.me/91{re_target_phone}?text={urllib.parse.quote(wa_re_msg)}"
                col_wa2.markdown(f'<a href="{re_wa_url}" target="_blank"><button style="width:100%; height:45px; background-color:#25D366; color:white; border:none; border-radius:8px; font-weight:bold; font-size:15px; cursor:pointer;">📲 Send Invoice via WhatsApp</button></a>', unsafe_allow_html=True)
        else:
            st.info("No previous invoices archived yet.")

    # ----------------- MODE C: EXPENSE VOUCHER -----------------
    elif bill_type == "Payment Voucher (Expense)":
        col_in1, col_in2 = st.columns(2)
        bill_no = col_in1.text_input("Voucher No.", f"VOU-{datetime.now().strftime('%Y%m%d%H%M')}")
        bill_date = col_in2.date_input("Voucher Date", datetime.now()).strftime("%Y-%m-%d")
        
        cust_name = col_in1.text_input("Paid To / Vendor Name *")
        cust_phone = col_in2.text_input("Contact Mobile (Optional)", "")
        
        col_e1, col_e2 = st.columns(2)
        exp_serv_sel = col_e1.selectbox("Expense Category / Purpose *", [
            "Office Rent", "Electricity & Utility", "Stationary & Printing", 
            "Software & Internet", "Staff Salary / Tea & Refreshment", "Government Fees / Challan", "OTHER"
        ])
        service_1 = col_e1.text_input("Type Custom Expense Details *") if exp_serv_sel == "OTHER" else exp_serv_sel
        amt_1 = col_e2.number_input("Amount (₹) *", min_value=0.0, step=50.0)
        pay_mode = col_in1.selectbox("Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"])
        remarks = col_in2.text_input("Remarks / Notes", "Office Accounting Voucher")
        auto_save = st.checkbox("☑ Automatically post this entry to Expenses", value=True)

        if st.button("💾 Generate Expense Voucher & Auto Post", type="primary", use_container_width=True):
            if cust_name and amt_1 > 0 and service_1:
                if auto_save:
                    ExcelManager.append_row("Expense", {
                        "Date": bill_date,
                        "Expense Name": f"{cust_name} ({service_1})",
                        "Amount": amt_1,
                        "Notes": f"Voucher #{bill_no} | {pay_mode} | {remarks}"
                    })
                    st.success("✅ Expense posted successfully!")
            else:
                st.error("Vendor name, service details, and amount are required.")

    # ----------------- MODE D: DIRECT DUE SETTLEMENT -----------------
    elif bill_type == "Settle Old Pending Due (Direct Payment)":
        st.markdown("### 💵 Direct Settle Customer Pending Balance")
        df_baki = ExcelManager.get_df("Udhar_Baki")
        
        if not df_baki.empty:
            pending_dues = df_baki[df_baki["Pending Amount"] > 0]
            if not pending_dues.empty:
                serv_c = "Service Details" if "Service Details" in df_baki else "Primary Service / Purpose"
                due_options = [f"ID #{r['ID']} - {r['Customer Name']} ({r['Mobile Number']}) | Item: {r.get(serv_c, 'Service')} | Pending: ₹{r['Pending Amount']:,.2f}" for _, r in pending_dues.iterrows()]
                selected_due = st.selectbox("Select Pending Customer Account to Settle:", due_options)
                
                sel_due_id = int(selected_due.split(" ")[1].replace("#", ""))
                due_row = df_baki[df_baki["ID"] == sel_due_id].iloc[0]
                
                c_name = due_row["Customer Name"]
                c_phone = due_row["Mobile Number"]
                c_serv = due_row.get(serv_c, "Service")
                curr_pending = float(due_row["Pending Amount"])
                
                st.info(f"👤 **Customer:** {c_name} | 📞 **Mobile:** {c_phone} | 🏷️ **Pending For:** {c_serv} | **Total Due:** ₹{curr_pending:,.2f}")
                
                col_d1, col_d2 = st.columns(2)
                settle_date = col_d1.date_input("Settlement Date", datetime.now()).strftime("%Y-%m-%d")
                settle_amt = col_d2.number_input("Deposit / Payment Received Now (₹) *", min_value=0.0, max_value=curr_pending, value=curr_pending, step=100.0)
                settle_mode = col_d1.selectbox("Deposit Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"])
                settle_notes = col_d2.text_input("Settlement Note", f"Settlement against {c_serv} (Due Rec #{sel_due_id})")
                
                if st.button("💳 Settle Balance & Record as Income", type="primary", use_container_width=True):
                    if settle_amt > 0:
                        new_paid = float(due_row["Paid Amount"]) + settle_amt
                        new_pending = curr_pending - settle_amt
                        new_stat = "Cleared" if new_pending <= 0 else "Pending"
                        
                        ExcelManager.update_row("Udhar_Baki", sel_due_id, {
                            "Paid Amount": new_paid,
                            "Pending Amount": new_pending,
                            "Status": new_stat
                        })
                        
                        ExcelManager.append_row("Income", {
                            "Date": settle_date,
                            "Customer/Person": c_name,
                            "Work Details": f"Due Settlement for {c_serv} (Rec #{sel_due_id})",
                            "Amount": settle_amt,
                            "Payment Mode": settle_mode,
                            "Notes": f"Mob: {c_phone} | {settle_notes}"
                        })
                        
                        st.success(f"✅ Successfully settled ₹{settle_amt:,.2f}! Remaining balance: ₹{new_pending:,.2f}")
                        st.rerun()
            else:
                st.success("No pending dues found in records!")
        else:
            st.info("No credit ledger records found.")

# ----------------- 4. REPORTS & PDF MODULE (ALL IN ONE + SEPARATE REPORTS) -----------------
elif menu == "📄 Reports & PDF":
    st.subheader("📄 Financial Statements & Periodic Reports")
    
    col_r1, col_r2 = st.columns(2)
    d_from = col_r1.date_input("From Date", datetime.now().replace(day=1)).strftime("%Y-%m-%d")
    d_to = col_r2.date_input("To Date", datetime.now()).strftime("%Y-%m-%d")

    df_i = ExcelManager.get_df("Income")
    df_e = ExcelManager.get_df("Expense")
    df_b = ExcelManager.get_df("Udhar_Baki")

    inc_date_col = "Date" if "Date" in df_i else None
    inc_amt_col = "Amount" if "Amount" in df_i else None
    exp_date_col = "Date" if "Date" in df_e else None
    exp_amt_col = "Amount" if "Amount" in df_e else None
    baki_date_col = "Date" if "Date" in df_b else None

    f_i = df_i[(df_i[inc_date_col] >= d_from) & (df_i[inc_date_col] <= d_to)] if not df_i.empty and inc_date_col else pd.DataFrame()
    f_e = df_e[(df_e[exp_date_col] >= d_from) & (df_e[exp_date_col] <= d_to)] if not df_e.empty and exp_date_col else pd.DataFrame()
    f_b = df_b[(df_b[baki_date_col] >= d_from) & (df_b[baki_date_col] <= d_to)] if not df_b.empty and baki_date_col else (df_b if not df_b.empty else pd.DataFrame())

    t_i = f_i[inc_amt_col].sum() if not f_i.empty and inc_amt_col else 0
    t_e = f_e[exp_amt_col].sum() if not f_e.empty and exp_amt_col else 0
    t_due = f_b["Pending Amount"].sum() if not f_b.empty and "Pending Amount" in f_b else 0
    
    cash_op, bank_op = ExcelManager.get_opening_balance()
    tot_op = cash_op + bank_op
    closing_bal = tot_op + t_i - t_e

    st.info(f"**Period:** {d_from} to {d_to} | **Opening Balance:** ₹{tot_op:,.2f} | **Total Revenue:** ₹{t_i:,.2f} | **Total Expenses:** ₹{t_e:,.2f} | **Closing Balance:** ₹{closing_bal:,.2f} | **Period Dues:** ₹{t_due:,.2f}")

    tab_all_rep, tab_inc_rep, tab_exp_rep, tab_due_rep = st.tabs([
        "📊 Complete Financial Statement",
        "💰 Income Only Report",
        "💸 Expense Only Report",
        "📋 Due / Outstanding Collection Report"
    ])

    # 1. Complete Statement
    with tab_all_rep:
        st.markdown("##### 📊 Comprehensive Statement (Revenue + Expenses + Closing Balance)")
        c_i1, c_i2 = st.columns(2)
        with c_i1:
            st.markdown("###### 💰 Income Breakdown")
            st.dataframe(f_i, use_container_width=True)
        with c_i2:
            st.markdown("###### 💸 Expense Breakdown")
            st.dataframe(f_e, use_container_width=True)
            
        def get_comprehensive_pdf():
            buf = io.BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=20, bottomMargin=20)
            elems = []
            styles = getSampleStyleSheet()
            
            c_title = ParagraphStyle('T1', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor("#1E3A8A"), alignment=1)
            c_sub = ParagraphStyle('T2', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor("#475569"), alignment=1)
            
            logo_l = PDFImage(LOGO_VISA, width=50, height=50) if os.path.exists(LOGO_VISA) else ""
            logo_r = PDFImage(LOGO_FINCARE, width=75, height=38) if os.path.exists(LOGO_FINCARE) else ""
                
            hdr_table_data = [[logo_l, [Paragraph(f"<b>{COMPANY_NAME}</b>", c_title), Paragraph(f"📍 {COMPANY_ADDRESS} | 📞 Phone: {COMPANY_MOBILE}", c_sub), Paragraph(f"COMPREHENSIVE FINANCIAL STATEMENT ({d_from} to {d_to})", ParagraphStyle('T3', parent=styles['Heading2'], fontSize=11, alignment=1))], logo_r]]
            t_hdr = Table(hdr_table_data, colWidths=[65, 415, 75])
            t_hdr.setStyle(TableStyle([('ALIGN', (0, 0), (0, -1), 'LEFT'), ('ALIGN', (1, 0), (1, -1), 'CENTER'), ('ALIGN', (2, 0), (2, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
            elems.append(t_hdr)
            elems.append(Spacer(1, 8))
            elems.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceAfter=8))

            sum_tbl = Table([
                ["Opening Balance", f"Rs. {tot_op:,.2f}"],
                ["Total Revenue", f"Rs. {t_i:,.2f}"],
                ["Total Expenses", f"Rs. {t_e:,.2f}"],
                ["Closing Balance", f"Rs. {closing_bal:,.2f}"]
            ], colWidths=[200, 200])
            sum_tbl.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F8FAFC")), ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E1"))]))
            elems.append(sum_tbl)
            elems.append(Spacer(1, 12))

            if not f_i.empty:
                elems.append(Paragraph("<b>Revenue Details:</b>", styles['Heading3']))
                i_rows = [["Date", "Customer / Party", "Description", "Mode", "Amount (Rs.)"]]
                for _, r in f_i.iterrows():
                    i_rows.append([str(r[inc_date_col]), str(r.get("Customer/Person", "-")), str(r.get("Work Details", "-")), str(r.get("Payment Mode", "-")), f"{float(r[inc_amt_col]):,.2f}"])
                t1 = Table(i_rows, colWidths=[65, 120, 160, 85, 100])
                t1.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E3A8A")), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1"))]))
                elems.append(t1)
                elems.append(Spacer(1, 12))

            if not f_e.empty:
                elems.append(Paragraph("<b>Expense Details:</b>", styles['Heading3']))
                e_rows = [["Date", "Expense Particulars", "Notes", "Amount (Rs.)"]]
                for _, r in f_e.iterrows():
                    e_rows.append([str(r[exp_date_col]), str(r.get("Expense Name", "-")), str(r.get("Notes", "-")), f"{float(r[exp_amt_col]):,.2f}"])
                t2 = Table(e_rows, colWidths=[75, 230, 115, 110])
                t2.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#DC2626")), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1"))]))
                elems.append(t2)

            doc.build(elems)
            buf.seek(0)
            return buf
            
        st.download_button("📥 Download Comprehensive Financial Statement PDF", data=get_comprehensive_pdf(), file_name=f"Comprehensive_Statement_{d_from}_to_{d_to}.pdf", mime="application/pdf", use_container_width=True)

    # 2. Income Only Report
    with tab_inc_rep:
        st.markdown(f"##### 💰 Income / Revenue Statement ({d_from} to {d_to})")
        st.write(f"**Total Revenue in Period:** ₹ {t_i:,.2f}")
        st.dataframe(f_i, use_container_width=True)
        
        def get_income_pdf():
            buf = io.BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=20, bottomMargin=20)
            elems = []
            styles = getSampleStyleSheet()
            c_title = ParagraphStyle('T1', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor("#1E3A8A"), alignment=1)
            c_sub = ParagraphStyle('T2', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor("#475569"), alignment=1)
            logo_l = PDFImage(LOGO_VISA, width=50, height=50) if os.path.exists(LOGO_VISA) else ""
            logo_r = PDFImage(LOGO_FINCARE, width=75, height=38) if os.path.exists(LOGO_FINCARE) else ""
            
            hdr_table_data = [[logo_l, [Paragraph(f"<b>{COMPANY_NAME}</b>", c_title), Paragraph(f"📍 {COMPANY_ADDRESS} | 📞 Phone: {COMPANY_MOBILE}", c_sub), Paragraph(f"INCOME & REVENUE REPORT ({d_from} to {d_to})", ParagraphStyle('T3', parent=styles['Heading2'], fontSize=11, alignment=1))], logo_r]]
            t_hdr = Table(hdr_table_data, colWidths=[65, 415, 75])
            t_hdr.setStyle(TableStyle([('ALIGN', (0, 0), (0, -1), 'LEFT'), ('ALIGN', (1, 0), (1, -1), 'CENTER'), ('ALIGN', (2, 0), (2, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
            elems.append(t_hdr)
            elems.append(Spacer(1, 8))
            elems.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceAfter=8))
            
            sum_tbl = Table([["Total Revenue", f"Rs. {t_i:,.2f}"], ["Total Records", f"{len(f_i)} Entries"]], colWidths=[200, 200])
            sum_tbl.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F0FDF4")), ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E1"))]))
            elems.append(sum_tbl)
            elems.append(Spacer(1, 12))

            if not f_i.empty:
                i_rows = [["Date", "Customer / Party", "Description", "Mode", "Amount (Rs.)"]]
                for _, r in f_i.iterrows():
                    i_rows.append([str(r[inc_date_col]), str(r.get("Customer/Person", "-")), str(r.get("Work Details", "-")), str(r.get("Payment Mode", "-")), f"{float(r[inc_amt_col]):,.2f}"])
                t1 = Table(i_rows, colWidths=[65, 120, 160, 85, 100])
                t1.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#15803D")), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1"))]))
                elems.append(t1)
            doc.build(elems)
            buf.seek(0)
            return buf
            
        st.download_button("📥 Download Income Report PDF", data=get_income_pdf(), file_name=f"Income_Report_{d_from}_to_{d_to}.pdf", mime="application/pdf", use_container_width=True)

    # 3. Expense Only Report
    with tab_exp_rep:
        st.markdown(f"##### 💸 Expense Statement ({d_from} to {d_to})")
        st.write(f"**Total Expenses in Period:** ₹ {t_e:,.2f}")
        st.dataframe(f_e, use_container_width=True)
        
        def get_expense_pdf():
            buf = io.BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=20, bottomMargin=20)
            elems = []
            styles = getSampleStyleSheet()
            c_title = ParagraphStyle('T1', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor("#1E3A8A"), alignment=1)
            c_sub = ParagraphStyle('T2', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor("#475569"), alignment=1)
            logo_l = PDFImage(LOGO_VISA, width=50, height=50) if os.path.exists(LOGO_VISA) else ""
            logo_r = PDFImage(LOGO_FINCARE, width=75, height=38) if os.path.exists(LOGO_FINCARE) else ""
            
            hdr_table_data = [[logo_l, [Paragraph(f"<b>{COMPANY_NAME}</b>", c_title), Paragraph(f"📍 {COMPANY_ADDRESS} | 📞 Phone: {COMPANY_MOBILE}", c_sub), Paragraph(f"EXPENSE STATEMENT REPORT ({d_from} to {d_to})", ParagraphStyle('T3', parent=styles['Heading2'], fontSize=11, alignment=1))], logo_r]]
            t_hdr = Table(hdr_table_data, colWidths=[65, 415, 75])
            t_hdr.setStyle(TableStyle([('ALIGN', (0, 0), (0, -1), 'LEFT'), ('ALIGN', (1, 0), (1, -1), 'CENTER'), ('ALIGN', (2, 0), (2, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
            elems.append(t_hdr)
            elems.append(Spacer(1, 8))
            elems.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceAfter=8))
            
            sum_tbl = Table([["Total Expenses", f"Rs. {t_e:,.2f}"], ["Total Records", f"{len(f_e)} Entries"]], colWidths=[200, 200])
            sum_tbl.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#FEF2F2")), ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E1"))]))
            elems.append(sum_tbl)
            elems.append(Spacer(1, 12))

            if not f_e.empty:
                e_rows = [["Date", "Expense Particulars", "Notes", "Amount (Rs.)"]]
                for _, r in f_e.iterrows():
                    e_rows.append([str(r[exp_date_col]), str(r.get("Expense Name", "-")), str(r.get("Notes", "-")), f"{float(r[exp_amt_col]):,.2f}"])
                t2 = Table(e_rows, colWidths=[75, 230, 115, 110])
                t2.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#DC2626")), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1"))]))
                elems.append(t2)
            doc.build(elems)
            buf.seek(0)
            return buf
            
        st.download_button("📥 Download Expense Report PDF", data=get_expense_pdf(), file_name=f"Expense_Report_{d_from}_to_{d_to}.pdf", mime="application/pdf", use_container_width=True)

    # 4. Due / Outstanding Collection Report
    with tab_due_rep:
        st.markdown(f"##### 📋 Outstanding Due & Collections Statement")
        active_pending_dues = f_b[f_b["Pending Amount"] > 0] if not f_b.empty and "Pending Amount" in f_b else pd.DataFrame()
        total_active_due_sum = active_pending_dues["Pending Amount"].sum() if not active_pending_dues.empty else 0.0
        
        st.write(f"**Total Outstanding Pending Due:** ₹ {total_active_due_sum:,.2f} ({len(active_pending_dues)} Pending Accounts)")
        st.dataframe(f_b, use_container_width=True)
        
        def get_due_report_pdf():
            buf = io.BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=20, bottomMargin=20)
            elems = []
            styles = getSampleStyleSheet()
            c_title = ParagraphStyle('T1', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor("#1E3A8A"), alignment=1)
            c_sub = ParagraphStyle('T2', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor("#475569"), alignment=1)
            logo_l = PDFImage(LOGO_VISA, width=50, height=50) if os.path.exists(LOGO_VISA) else ""
            logo_r = PDFImage(LOGO_FINCARE, width=75, height=38) if os.path.exists(LOGO_FINCARE) else ""
            
            hdr_table_data = [[logo_l, [Paragraph(f"<b>{COMPANY_NAME}</b>", c_title), Paragraph(f"📍 {COMPANY_ADDRESS} | 📞 Phone: {COMPANY_MOBILE}", c_sub), Paragraph("OUTSTANDING DUE & CREDIT COLLECTION REPORT", ParagraphStyle('T3', parent=styles['Heading2'], fontSize=11, alignment=1))], logo_r]]
            t_hdr = Table(hdr_table_data, colWidths=[65, 415, 75])
            t_hdr.setStyle(TableStyle([('ALIGN', (0, 0), (0, -1), 'LEFT'), ('ALIGN', (1, 0), (1, -1), 'CENTER'), ('ALIGN', (2, 0), (2, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
            elems.append(t_hdr)
            elems.append(Spacer(1, 8))
            elems.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1E3A8A"), spaceAfter=8))
            
            sum_tbl = Table([["Total Active Pending Dues", f"Rs. {total_active_due_sum:,.2f}"], ["Pending Accounts Count", f"{len(active_pending_dues)} Clients"]], colWidths=[240, 200])
            sum_tbl.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#FFFBEB")), ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E1"))]))
            elems.append(sum_tbl)
            elems.append(Spacer(1, 12))

            if not f_b.empty:
                d_rows = [["Date", "Customer Name", "Mobile", "Service/Item", "Total (Rs.)", "Paid (Rs.)", "Pending (Rs.)", "Status"]]
                for _, r in f_b.iterrows():
                    d_rows.append([
                        str(r.get("Date", "-")),
                        str(r.get("Customer Name", "-")),
                        str(r.get("Mobile Number", "-")),
                        str(r.get("Service Details", "Service")),
                        f"{float(r.get('Total Amount', 0)):,.2f}",
                        f"{float(r.get('Paid Amount', 0)):,.2f}",
                        f"{float(r.get('Pending Amount', 0)):,.2f}",
                        str(r.get("Status", "Pending"))
                    ])
                t3 = Table(d_rows, colWidths=[55, 95, 75, 90, 60, 55, 65, 45])
                t3.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#D97706")), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")), ('FONTSIZE', (0, 0), (-1, -1), 8)]))
                elems.append(t3)
            doc.build(elems)
            buf.seek(0)
            return buf
            
        st.download_button("📥 Download Outstanding Due Report PDF", data=get_due_report_pdf(), file_name=f"Due_Collections_Report_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf", use_container_width=True)

# ----------------- 5. OPENING BALANCE (CHOPADA MIGRATION) -----------------
elif menu == "🏦 Opening Balance":
    st.subheader("🏦 Opening Balance Setup & Ledger Migration")
    st.info("💡 Transfer your existing book balances (Cash in Hand, Bank Balance, and Previous Customer Dues) directly into this software.")
    
    tab_cash_bank, tab_cust_op = st.tabs(["💵 Cash & Bank Opening Balance", "👥 Transfer Customer Previous Dues"])
    
    with tab_cash_bank:
        current_cash, current_bank = ExcelManager.get_opening_balance()
        st.markdown("##### 💼 Set Initial Cash on Hand & Bank Account Balance")
        
        col_b1, col_b2 = st.columns(2)
        in_cash_op = col_b1.number_input("Opening Cash in Hand (₹) *", min_value=0.0, value=float(current_cash), step=500.0)
        in_bank_op = col_b2.number_input("Opening Bank Account Balance (₹) *", min_value=0.0, value=float(current_bank), step=500.0)
        
        total_op_val = in_cash_op + in_bank_op
        st.write(f"#### **Total Initial Starting Capital: ₹ {total_op_val:,.2f}**")
        
        st.markdown("🔒 **Security PIN Required to Update Opening Balance:**")
        op_pin = st.text_input("Enter Security PIN:", type="password", key="op_pin_cash")
        
        if st.button("💾 Save Cash & Bank Opening Balance", type="primary", use_container_width=True):
            if op_pin == get_saved_pin():
                ExcelManager.set_opening_balance(in_cash_op, in_bank_op)
                st.success("✅ Opening Cash & Bank balances successfully saved! Dashboard updated.")
                st.rerun()
            else:
                st.error("❌ Incorrect Security PIN! Action denied.")

    with tab_cust_op:
        st.markdown("##### 👥 Direct Transfer of Previous Customer Outstandings")
        st.caption("Enter customers who have pending dues in your previous books with specific service details.")
        
        with st.form("op_baki_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            op_cname = c1.text_input("Customer Name *")
            op_cphone = c2.text_input("Mobile Number (10 Digits) *")
            
            c3, c4 = st.columns(2)
            op_due_amt = c3.number_input("Previous Pending Due Amount (₹) *", min_value=0.0, step=100.0)
            op_service_sel = c4.selectbox("Pending For Which Service / Item *", SERVICE_OPTIONS)
            op_service = st.text_input("Type Custom Service Name *") if op_service_sel == "OTHER" else op_service_sel
            
            c5, c6 = st.columns(2)
            op_date = c5.date_input("Due / Entry Date", datetime.now()).strftime("%Y-%m-%d")
            op_note = c6.text_input("Notes", "Brought forward from previous books")
            
            st.markdown("🔒 **Security PIN:**")
            cust_op_pin = st.text_input("Enter Security PIN to Confirm:", type="password", key="op_pin_cust")
            
            if st.form_submit_button("💾 Migrate Customer & Add to Due Ledger", use_container_width=True):
                if cust_op_pin == get_saved_pin():
                    if op_cname and op_cphone and op_due_amt > 0 and op_service:
                        clean_p = str(op_cphone).strip()
                        df_c = ExcelManager.get_df("Customers")
                        if df_c.empty or "Mobile Number" not in df_c or clean_p not in df_c["Mobile Number"].astype(str).values:
                            ExcelManager.append_row("Customers", {
                                "Created Date": op_date,
                                "Customer Name": op_cname,
                                "Mobile Number": clean_p,
                                "City/Address": "Kadi",
                                "Primary Service / Purpose": op_service,
                                "Notes": op_note
                            })
                        
                        ExcelManager.append_row("Udhar_Baki", {
                            "Date": op_date,
                            "Customer Name": op_cname,
                            "Mobile Number": clean_p,
                            "Service Details": op_service,
                            "Total Amount": op_due_amt,
                            "Paid Amount": 0.0,
                            "Pending Amount": op_due_amt,
                            "Due Date": op_date,
                            "Status": "Pending"
                        })
                        st.success(f"✅ Successfully added Opening Due of ₹{op_due_amt:,.2f} for {op_cname} ({op_service})!")
                        st.rerun()
                    else:
                        st.error("Customer name, mobile number, and valid due amount are required.")
                else:
                    st.error("❌ Incorrect Security PIN!")

# ----------------- 6. CUSTOMERS DIRECTORY & PROMOTIONS -----------------
elif menu == "👥 Customers Directory":
    st.subheader("👥 Customer Directory & Bulk Broadcast Promotions")
    
    tab_new_cust, tab_list_cust, tab_promo = st.tabs(["➕ Add New Customer", "📋 Registered Clients List & Edit (PIN Protected)", "📢 Bulk Promotion / Offer Broadcast"])
    
    with tab_new_cust:
        with st.form("cust_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            c_name = c1.text_input("Customer Name *")
            c_phone = c2.text_input("Mobile Number (10 Digits) *")
            c_city = c1.text_input("City / Village / Address", "Kadi")
            
            c_service_sel = c2.selectbox("Primary Purpose / Came for Service *", SERVICE_OPTIONS)
            c_service = st.text_input("Type Custom Service Name *") if c_service_sel == "OTHER" else c_service_sel
            c_notes = st.text_area("Additional Requirements / Notes", placeholder="e.g. Canada Visa inquiry, Land valuation, etc.")
            
            if st.form_submit_button("💾 Save Customer Profile", use_container_width=True):
                if c_name and c_phone and c_service:
                    df_c = ExcelManager.get_df("Customers")
                    clean_phone = str(c_phone).strip()
                    if not df_c.empty and "Mobile Number" in df_c and clean_phone in df_c["Mobile Number"].astype(str).values:
                        st.warning(f"⚠️ A customer with mobile {clean_phone} already exists in records!")
                    else:
                        ExcelManager.append_row("Customers", {
                            "Created Date": datetime.now().strftime("%Y-%m-%d"),
                            "Customer Name": c_name,
                            "Mobile Number": clean_phone,
                            "City/Address": c_city,
                            "Primary Service / Purpose": c_service,
                            "Notes": c_notes
                        })
                        st.success(f"Customer '{c_name}' successfully added to directory!")
                        st.rerun()
                else:
                    st.error("Customer name, mobile number, and valid service are mandatory.")

    with tab_list_cust:
        df_c = ExcelManager.get_df("Customers")
        if not df_c.empty:
            search_query = st.text_input("🔍 Quick Search by Name or Mobile:", "")
            filtered_df = df_c[df_c["Customer Name"].astype(str).str.contains(search_query, case=False, na=False) | 
                               df_c["Mobile Number"].astype(str).str.contains(search_query, case=False, na=False)] if search_query else df_c
                
            st.dataframe(filtered_df, use_container_width=True)
            st.divider()
            
            sel_c_id = st.selectbox("Select Customer ID to Edit / Delete:", df_c["ID"].tolist())
            c_row = df_c[df_c["ID"] == sel_c_id].iloc[0]
            
            with st.expander(f"🔐 Edit Profile #{sel_c_id} - {c_row.get('Customer Name')} (Requires PIN)", expanded=True):
                u_cname = st.text_input("Name", str(c_row.get("Customer Name", "")))
                u_cphone = st.text_input("Mobile", str(c_row.get("Mobile Number", "")))
                u_ccity = st.text_input("City/Address", str(c_row.get("City/Address", "")))
                
                current_serv = str(c_row.get("Primary Service / Purpose", "OTHER"))
                serv_index = SERVICE_OPTIONS.index(current_serv) if current_serv in SERVICE_OPTIONS else SERVICE_OPTIONS.index("OTHER")
                u_cserv_sel = st.selectbox("Primary Service", SERVICE_OPTIONS, index=serv_index)
                u_cserv = st.text_input("Custom Service Name", value="" if current_serv in SERVICE_OPTIONS else current_serv) if u_cserv_sel == "OTHER" else u_cserv_sel
                u_cnotes = st.text_area("Notes", str(c_row.get("Notes", "")) if not pd.isna(c_row.get("Notes")) else "")
                
                st.markdown("🔒 **Security Verification:**")
                edit_pin = st.text_input("Enter PIN to Authorize Changes:", type="password", key="c_pin")
                
                b1, b2 = st.columns(2)
                if b1.button("🔄 Update Customer", use_container_width=True):
                    if edit_pin == get_saved_pin():
                        ExcelManager.update_row("Customers", sel_c_id, {
                            "Customer Name": u_cname,
                            "Mobile Number": u_cphone,
                            "City/Address": u_ccity,
                            "Primary Service / Purpose": u_cserv,
                            "Notes": u_cnotes
                        })
                        st.success("Customer profile updated!")
                        st.rerun()
                    else:
                        st.error("❌ Incorrect Security PIN! Action denied.")
                        
                if b2.button("🗑️ Delete Customer", type="primary", use_container_width=True):
                    if edit_pin == get_saved_pin():
                        ExcelManager.delete_row("Customers", sel_c_id)
                        st.warning("Customer deleted from directory!")
                        st.rerun()
                    else:
                        st.error("❌ Incorrect Security PIN! Action denied.")
        else:
            st.info("No registered clients found.")

    with tab_promo:
        st.markdown("##### 📢 Bulk Promotion & Offer Broadcast System")
        st.info("💡 You can send customized festival offers, visa updates, or schemes to all customers at once.")
        
        df_c = ExcelManager.get_df("Customers")
        if not df_c.empty:
            col_f1, col_f2 = st.columns([1.5, 2])
            target_audience = col_f1.selectbox("Select Target Audience:", ["All Registered Customers"] + list(df_c["Primary Service / Purpose"].dropna().unique()))
            target_df = df_c if target_audience == "All Registered Customers" else df_c[df_c["Primary Service / Purpose"] == target_audience]
            
            col_f2.metric("Total Selected Recipients", f"{len(target_df)} Clients")
            
            st.markdown("##### ✍️ Compose Broadcast Message:")
            st.caption("Tags available: `{name}` for Customer Name | `{service}` for Service Category")
            
            default_template = f"Hello {{name}}, special update and exclusive offer from {COMPANY_NAME}! For any inquiries regarding {{service}}, please contact us today. Address: {COMPANY_ADDRESS} | Phone: {COMPANY_MOBILE}"
            broadcast_msg = st.text_area("Message Content:", value=default_template, height=100)
            
            all_mobiles_str = ", ".join(target_df["Mobile Number"].dropna().astype(str).tolist())
            st.text_area("📋 Copy All Numbers (For WhatsApp Broadcast Lists):", value=all_mobiles_str, height=65)
            
            st.divider()
            st.markdown("##### 📲 One-Click Personalized WhatsApp Trigger List:")
            
            for idx, prow in target_df.iterrows():
                c_n = str(prow.get("Customer Name", "Customer"))
                c_p = str(prow.get("Mobile Number", "")).strip()
                c_s = str(prow.get("Primary Service / Purpose", "Service"))
                
                cust_msg = broadcast_msg.replace("{name}", c_n).replace("{service}", c_s)
                p_url = f"https://wa.me/91{c_p}?text={urllib.parse.quote(cust_msg)}"
                
                cols_b = st.columns([2.5, 2, 2, 2])
                cols_b[0].write(f"**{c_n}**")
                cols_b[1].write(f"📞 {c_p}")
                cols_b[2].write(f"🏷️ {c_s}")
                cols_b[3].markdown(f"[📲 Send to {c_n.split()[0]}]({p_url})", unsafe_allow_html=True)
        else:
            st.info("No customer directory data available.")

# ----------------- 7. INCOME (PIN PROTECTED EDIT/DELETE) -----------------
elif menu == "💰 Income":
    st.subheader("💰 Income Records & Management")
    st.info("💡 Note: Income entries are automatically generated when creating invoices or settling customer dues.")
    
    df_inc = ExcelManager.get_df("Income")
    if not df_inc.empty:
        st.dataframe(df_inc, use_container_width=True)
        st.divider()
        sel_id = st.selectbox("Select Record ID to Edit / Delete:", df_inc["ID"].tolist())
        row = df_inc[df_inc["ID"] == sel_id].iloc[0]
        
        d_k = "Date"
        n_k = "Customer/Person"
        w_k = "Work Details"
        a_k = "Amount"
        p_k = "Payment Mode"
        not_k = "Notes"
        
        with st.expander(f"🔐 Edit Income Record ID #{sel_id} (Requires PIN)", expanded=True):
            u_date = st.text_input("Date", str(row.get(d_k, "")))
            u_name = st.text_input("Customer Name", str(row.get(n_k, "")))
            u_work = st.text_input("Work Details", str(row.get(w_k, "")))
            u_amt = st.number_input("Amount (₹)", value=float(row.get(a_k, 0.0)), step=100.0)
            current_mode = str(row.get(p_k, "Cash"))
            modes_list = ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"]
            u_mode = st.selectbox("Payment Mode", modes_list, index=modes_list.index(current_mode) if current_mode in modes_list else 0)
            u_note = st.text_input("Notes", str(row.get(not_k, "")) if not pd.isna(row.get(not_k)) else "")
            
            st.markdown("🔒 **Security Verification:**")
            inc_pin = st.text_input("Enter Security PIN to Authorize:", type="password", key="inc_pin")
            
            b_up, b_del = st.columns(2)
            if b_up.button("🔄 Update Record", use_container_width=True):
                if inc_pin == get_saved_pin():
                    ExcelManager.update_row("Income", sel_id, {
                        "Date": u_date, "Customer/Person": u_name, "Work Details": u_work,
                        "Amount": u_amt, "Payment Mode": u_mode, "Notes": u_note
                    })
                    st.success("Record updated successfully!")
                    st.rerun()
                else:
                    st.error("❌ Incorrect PIN! Action denied.")
                    
            if b_del.button("🗑️ Delete Record", type="primary", use_container_width=True):
                if inc_pin == get_saved_pin():
                    ExcelManager.delete_row("Income", sel_id)
                    st.warning("Record deleted!")
                    st.rerun()
                else:
                    st.error("❌ Incorrect PIN! Action denied.")
    else:
        st.info("No income records found.")

# ----------------- 8. EXPENSE (PIN PROTECTED EDIT/DELETE) -----------------
elif menu == "💸 Expenses":
    st.subheader("💸 Expense Records & Management")
    st.info("💡 Note: Expense entries are automatically generated when creating a payment voucher under **🧾 Generate Bill / Voucher**.")
    
    df_exp = ExcelManager.get_df("Expense")
    if not df_exp.empty:
        st.dataframe(df_exp, use_container_width=True)
        st.divider()
        sel_id = st.selectbox("Select Record ID to Edit / Delete:", df_exp["ID"].tolist(), key="exp_sel")
        row = df_exp[df_exp["ID"] == sel_id].iloc[0]
        
        d_k = "Date"
        n_k = "Expense Name"
        a_k = "Amount"
        not_k = "Notes"
        
        with st.expander(f"🔐 Edit Expense Record ID #{sel_id} (Requires PIN)", expanded=True):
            u_date = st.text_input("Date", str(row.get(d_k, "")), key="e_date")
            u_name = st.text_input("Expense Name", str(row.get(n_k, "")), key="e_name")
            u_amt = st.number_input("Amount (₹)", value=float(row.get(a_k, 0.0)), step=50.0, key="e_amt")
            u_note = st.text_input("Notes", str(row.get(not_k, "")) if not pd.isna(row.get(not_k)) else "", key="e_note")
            
            st.markdown("🔒 **Security Verification:**")
            exp_pin = st.text_input("Enter Security PIN to Authorize:", type="password", key="exp_pin")
            
            b_up, b_del = st.columns(2)
            if b_up.button("🔄 Update Record", key="e_up", use_container_width=True):
                if exp_pin == get_saved_pin():
                    ExcelManager.update_row("Expense", sel_id, {
                        "Date": u_date, "Expense Name": u_name, "Amount": u_amt, "Notes": u_note
                    })
                    st.success("Expense updated successfully!")
                    st.rerun()
                else:
                    st.error("❌ Incorrect PIN! Action denied.")
                    
            if b_del.button("🗑️ Delete Record", key="e_del", type="primary", use_container_width=True):
                if exp_pin == get_saved_pin():
                    ExcelManager.delete_row("Expense", sel_id)
                    st.warning("Expense deleted!")
                    st.rerun()
                else:
                    st.error("❌ Incorrect PIN! Action denied.")
    else:
        st.info("No expense records found.")

# ----------------- 9. UDHAR / BAKI (PIN PROTECTED EDIT/DELETE) -----------------
elif menu == "📋 Due Collections":
    st.subheader("📋 Due Collections & Itemized Credit Ledger")
    
    df_baki = ExcelManager.get_df("Udhar_Baki")
    if not df_baki.empty:
        st.dataframe(df_baki, use_container_width=True)
        st.divider()
        sel_id = st.selectbox("Select Record ID to Edit / Settle / Delete:", df_baki["ID"].tolist(), key="baki_sel")
        row = df_baki[df_baki["ID"] == sel_id].iloc[0]
        
        n_k = "Customer Name"
        m_k = "Mobile Number"
        serv_k = "Service Details"
        t_k = "Total Amount"
        p_k = "Paid Amount"
        d_k = "Due Date"
        
        with st.expander(f"🔐 Settle / Edit Due Record ID #{sel_id} (Requires PIN)", expanded=True):
            u_name = st.text_input("Customer Name", str(row.get(n_k, "")), key="b_name")
            u_phone = st.text_input("Mobile Number", str(row.get(m_k, "")), key="b_phone")
            u_serv = st.text_input("Service / Item Name", str(row.get(serv_k, "Service")), key="b_serv")
            u_tot = st.number_input("Total Bill (₹)", value=float(row.get(t_k, 0.0)), step=100.0, key="b_tot")
            u_rec = st.number_input("Paid Amount (₹)", value=float(row.get(p_k, 0.0)), step=100.0, key="b_rec")
            u_due = st.text_input("Due Date", str(row.get(d_k, "")), key="b_due")
            
            st.markdown("🔒 **Security Verification:**")
            due_pin = st.text_input("Enter Security PIN to Authorize:", type="password", key="due_pin")
            
            b_up, b_del = st.columns(2)
            if b_up.button("🔄 Update & Recalculate Balance", key="b_up", use_container_width=True):
                if due_pin == get_saved_pin():
                    u_baki = u_tot - u_rec
                    ExcelManager.update_row("Udhar_Baki", sel_id, {
                        "Customer Name": u_name, "Mobile Number": u_phone,
                        "Service Details": u_serv,
                        "Total Amount": u_tot, "Paid Amount": u_rec, "Pending Amount": u_baki,
                        "Due Date": u_due, "Status": "Cleared" if u_baki <= 0 else "Pending"
                    })
                    st.success(f"Updated! New pending balance: ₹{u_baki:,.2f}")
                    st.rerun()
                else:
                    st.error("❌ Incorrect PIN! Action denied.")
                    
            if b_del.button("🗑️ Delete Record", key="b_del", type="primary", use_container_width=True):
                if due_pin == get_saved_pin():
                    ExcelManager.delete_row("Udhar_Baki", sel_id)
                    st.warning("Record deleted!")
                    st.rerun()
                else:
                    st.error("❌ Incorrect PIN! Action denied.")
    else:
        st.info("No credit records found.")

# ----------------- 10. SECURITY & PIN SETTINGS -----------------
elif menu == "⚙️ Security / Change PIN":
    st.subheader("⚙️ Security Settings & Change Master PIN")
    st.info("🔒 This PIN protects your system login as well as all record modification and deletion actions.")
    
    with st.form("pin_change_form"):
        curr_pin_input = st.text_input("Enter Current PIN *", type="password")
        new_pin_input = st.text_input("Enter New PIN (4 to 8 Digits) *", type="password")
        conf_pin_input = st.text_input("Confirm New PIN *", type="password")
        
        if st.form_submit_button("💾 Update Security PIN", use_container_width=True):
            if curr_pin_input == get_saved_pin():
                if new_pin_input and new_pin_input == conf_pin_input:
                    save_new_pin(new_pin_input)
                    st.success("✅ Master Security PIN updated successfully!")
                else:
                    st.error("New PIN and Confirm PIN do not match.")
            else:
                st.error("❌ Current PIN is incorrect.")