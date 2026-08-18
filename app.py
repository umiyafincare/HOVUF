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

EXCEL_FILE = "Rojmed_Data.xlsx"
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
        div[data-testid="metric-container"] {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            padding: 14px 18px;
            border-radius: 10px;
            border-left: 4px solid #2563EB;
        }
        div[data-testid="stMetricValue"] {
            font-weight: 700;
            color: #0F172A;
        }
    </style>
""", unsafe_allow_html=True)

class DataManager:
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
            ws_set.append(["Master_PIN", DEFAULT_PIN, datetime.now().strftime("%Y-%m-%d")])
            wb.save(EXCEL_FILE)

    @staticmethod
    def get_df(sheet_name):
        DataManager.init_excel()
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
            return df.dropna(how="all") if df is not None else pd.DataFrame()
        except Exception:
            return pd.DataFrame()

    @staticmethod
    def save_df(sheet_name, df):
        DataManager.init_excel()
        try:
            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            st.error(f"Save error: {e}")

    @staticmethod
    def append_row(sheet_name, row_dict):
        df = DataManager.get_df(sheet_name)
        new_id = 1 if df.empty or "ID" not in df else (df["ID"].max() + 1 if not pd.isna(df["ID"].max()) else len(df) + 1)
        if "ID" in df.columns or sheet_name not in ["Invoices_Archive", "Settings"]:
            row_dict["ID"] = int(new_id)
        new_row_df = pd.DataFrame([row_dict])
        df = pd.concat([df, new_row_df], ignore_index=True)
        DataManager.save_df(sheet_name, df)
        return int(new_id)

    @staticmethod
    def delete_row(sheet_name, row_id):
        df = DataManager.get_df(sheet_name)
        if not df.empty and "ID" in df:
            df = df[df["ID"] != row_id]
            DataManager.save_df(sheet_name, df)

    @staticmethod
    def update_row(sheet_name, row_id, updated_dict):
        df = DataManager.get_df(sheet_name)
        if not df.empty and "ID" in df:
            idx = df.index[df["ID"] == row_id].tolist()
            if idx:
                for k, v in updated_dict.items():
                    df.at[idx[0], k] = v
                DataManager.save_df(sheet_name, df)

    @staticmethod
    def update_invoice_archive(invoice_no, updated_dict):
        df = DataManager.get_df("Invoices_Archive")
        if not df.empty and "Invoice No" in df:
            idx = df.index[df["Invoice No"].astype(str) == str(invoice_no)].tolist()
            if idx:
                for k, v in updated_dict.items():
                    df.at[idx[0], k] = v
                DataManager.save_df("Invoices_Archive", df)

    @staticmethod
    def delete_invoice_archive(invoice_no):
        df = DataManager.get_df("Invoices_Archive")
        if not df.empty and "Invoice No" in df:
            df = df[df["Invoice No"].astype(str) != str(invoice_no)]
            DataManager.save_df("Invoices_Archive", df)

    @staticmethod
    def get_pin():
        df_set = DataManager.get_df("Settings")
        if not df_set.empty and "Key" in df_set and "Value" in df_set:
            pin_row = df_set[df_set["Key"] == "Master_PIN"]
            if not pin_row.empty and pd.notna(pin_row["Value"].values[0]):
                return str(int(pin_row["Value"].values[0]) if isinstance(pin_row["Value"].values[0], (int, float)) else pin_row["Value"].values[0]).strip()
        return DEFAULT_PIN

    @staticmethod
    def set_pin(new_pin):
        df_set = DataManager.get_df("Settings")
        if df_set.empty or "Key" not in df_set:
            df_set = pd.DataFrame([{"Key": "Master_PIN", "Value": str(new_pin), "Updated Date": datetime.now().strftime("%Y-%m-%d")}])
        else:
            if "Master_PIN" in df_set["Key"].values:
                df_set.loc[df_set["Key"] == "Master_PIN", "Value"] = str(new_pin)
            else:
                df_set = pd.concat([df_set, pd.DataFrame([{"Key": "Master_PIN", "Value": str(new_pin), "Updated Date": datetime.now().strftime("%Y-%m-%d")}])], ignore_index=True)
        DataManager.save_df("Settings", df_set)

    @staticmethod
    def get_opening_balance():
        df_set = DataManager.get_df("Settings")
        cash_op, bank_op = 0.0, 0.0
        if not df_set.empty and "Key" in df_set and "Value" in df_set:
            cash_row = df_set[df_set["Key"] == "Cash_Opening_Balance"]
            bank_row = df_set[df_set["Key"] == "Bank_Opening_Balance"]
            if not cash_row.empty and pd.notna(cash_row["Value"].values[0]):
                cash_op = float(cash_row["Value"].values[0])
            if not bank_row.empty and pd.notna(bank_row["Value"].values[0]):
                bank_op = float(bank_row["Value"].values[0])
        return cash_op, bank_op

    @staticmethod
    def set_opening_balance(cash_op, bank_op):
        now_str = datetime.now().strftime("%Y-%m-%d")
        new_records = [
            {"Key": "Cash_Opening_Balance", "Value": float(cash_op), "Updated Date": now_str},
            {"Key": "Bank_Opening_Balance", "Value": float(bank_op), "Updated Date": now_str}
        ]
        df_set = DataManager.get_df("Settings")
        if df_set.empty or "Key" not in df_set:
            df_set = pd.DataFrame(new_records)
        else:
            for rec in new_records:
                if rec["Key"] in df_set["Key"].values:
                    df_set.loc[df_set["Key"] == rec["Key"], "Value"] = rec["Value"]
                else:
                    df_set = pd.concat([df_set, pd.DataFrame([rec])], ignore_index=True)
        DataManager.save_df("Settings", df_set)

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
                <h3 style="color: #1E3A8A; margin-top: 0;">🔒 Secure Ledger Login</h3>
                <p style="color: #64748B; font-size: 13px;">Enter 4-Digit Security PIN to access accounting ledger.</p>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        entered_pin = st.text_input("Enter 4-Digit Security PIN:", type="password", max_chars=8)
        
        if st.button("🔓 Unlock & Login", type="primary", use_container_width=True):
            if entered_pin == DataManager.get_pin():
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
    ("📄 Reports & PDF", "Statements & Reports"),
    ("🏦 Opening Balance", "Set Starting Balances"),
    ("👥 Customers Directory", "Manage Clients & Broadcasts"),
    ("💰 Income", "View & Manage Income"),
    ("💸 Expenses", "View & Manage Expenses"),
    ("📋 Due Collections", "Customer Pending Dues"),
    ("💾 Backup & Restore", "Download or Restore Previous Data"),
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

if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button("📥 Quick Backup (Excel)", data=f, file_name=f"Backup_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

if st.sidebar.button("🔒 Logout System", use_container_width=True):
    st.session_state.logged_in = False
    st.rerun()

SERVICE_OPTIONS = [
    "VISA", "PASSPORT", "PANCARD", "LOAN", "LAND RECORD (7/12 & 8-A)",
    "E-KYC ALL", "FARMER REGISTRATION", "PM-KISAN", "DIGITAL GUJARAT SERVICES",
    "PROVIDENT FUND (PF)", "REVENUE WORK", "MARRIAGE CERTIFICATE", "INSURANCE",
    "AIR TICKET", "OTHER"
]

def get_customer_pending_due(phone_number):
    if not phone_number:
        return 0.0, []
    df_baki = DataManager.get_df("Udhar_Baki")
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
    st.subheader("📊 Business Overview")
    
    df_rem_all = DataManager.get_df("Task_Reminder")
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
                    DataManager.update_row("Task_Reminder", t_id, {"Status": "Completed"})
                    st.success("Task Completed!")
                    st.rerun()
            st.divider()

    df_inc = DataManager.get_df("Income")
    df_exp = DataManager.get_df("Expense")
    df_baki = DataManager.get_df("Udhar_Baki")
    df_cust = DataManager.get_df("Customers")
    today_str = datetime.now().strftime("%Y-%m-%d")
    
    today_inc = df_inc[df_inc["Date"] == today_str]["Amount"].sum() if not df_inc.empty and "Date" in df_inc and "Amount" in df_inc else 0
    total_inc = df_inc["Amount"].sum() if not df_inc.empty and "Amount" in df_inc else 0
    today_exp = df_exp[df_exp["Date"] == today_str]["Amount"].sum() if not df_exp.empty and "Date" in df_exp and "Amount" in df_exp else 0
    total_exp = df_exp["Amount"].sum() if not df_exp.empty and "Amount" in df_exp else 0
    total_baki = df_baki["Pending Amount"].sum() if not df_baki.empty and "Pending Amount" in df_baki else 0
    
    cash_op, bank_op = DataManager.get_opening_balance()
    tot_op = cash_op + bank_op
    closing_net_balance = tot_op + total_inc - total_exp
    total_cust = len(df_cust) if not df_cust.empty else 0

    col0, col1, col2, col3, col4, col5 = st.columns(6)
    col0.metric("Opening Balance", f"₹ {tot_op:,.2f}", f"Cash: ₹{cash_op:,.0f} | Bank: ₹{bank_op:,.0f}")
    col1.metric("Today's Income", f"₹ {today_inc:,.2f}", f"Total: ₹ {total_inc:,.2f}")
    col2.metric("Today's Expense", f"₹ {today_exp:,.2f}", f"Total: ₹ {total_exp:,.2f}", delta_color="inverse")
    col3.metric("Closing Balance", f"₹ {closing_net_balance:,.2f}")
    col4.metric("Total Pending Dues", f"₹ {total_baki:,.2f}")
    col5.metric("Registered Clients", f"{total_cust}")

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
                    DataManager.append_row("Task_Reminder", {"Date": rdate, "Time": rtime, "Person Name": pname, "Mobile": rphone, "Task Details": tdesc, "Status": "Pending"})
                    st.success("Reminder Saved!")
                    st.rerun()

    with tab_pending_tasks:
        df_rem = DataManager.get_df("Task_Reminder")
        if not df_rem.empty:
            pending_list = df_rem[df_rem["Status"] == "Pending"]
            for _, r in pending_list.iterrows():
                r_id = r["ID"]
                c1, c2, c3, c4 = st.columns([2, 2, 3, 2])
                c1.write(f"📅 {r['Date']} | ⏰ {r['Time']}")
                c2.write(f"👤 {r.get('Person Name')}")
                c3.write(f"📌 {r.get('Task Details')}")
                if c4.button("✅ Done", key=f"done_{r_id}"):
                    DataManager.update_row("Task_Reminder", r_id, {"Status": "Completed"})
                    st.rerun()

    with tab_completed_tasks:
        df_rem = DataManager.get_df("Task_Reminder")
        if not df_rem.empty:
            st.dataframe(df_rem[df_rem["Status"] == "Completed"], use_container_width=True)

# ----------------- 3. INVOICE GENERATION & EDIT / DELETE WITH PIN -----------------
elif menu == "🧾 Generate Bill / Voucher":
    st.subheader("🧾 Generate, Edit & Manage Invoices / Vouchers")
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
        c1, c2 = st.columns(2)
        c_name = c1.text_input("Customer Name *")
        c_phone = c2.text_input("Mobile Number *")
        
        c3, c4 = st.columns(2)
        bill_no = c3.text_input("Invoice No.", f"INV-{datetime.now().strftime('%Y%m%d%H%M')}")
        bill_date = c4.date_input("Invoice Date", datetime.now()).strftime("%Y-%m-%d")
        
        c5, c6 = st.columns(2)
        s1_sel = c5.selectbox("Select Service 1 *", SERVICE_OPTIONS)
        s1 = c5.text_input("Custom Service Name *") if s1_sel == "OTHER" else s1_sel
        amt1 = c6.number_input("Amount 1 (₹) *", min_value=0.0, step=100.0)
        
        c7, c8 = st.columns(2)
        s2_sel = c7.selectbox("Service 2 (Optional)", ["None"] + SERVICE_OPTIONS)
        s2 = c7.text_input("Custom Service 2") if s2_sel == "OTHER" else ("" if s2_sel == "None" else s2_sel)
        amt2 = c8.number_input("Amount 2 (₹)", min_value=0.0, step=100.0)
        
        total_bill = amt1 + amt2
        st.write(f"### **Total Bill: ₹ {total_bill:,.2f}**")
        
        cp1, cp2, cp3 = st.columns(3)
        pay_mode = cp1.selectbox("Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque", "Pending / Due"])
        rec_amt = cp2.number_input("Received Amount (₹)", min_value=0.0, max_value=float(total_bill), value=float(total_bill) if pay_mode != "Pending / Due" else 0.0, step=100.0)
        due_date = cp3.date_input("Due Date", datetime.now()).strftime("%Y-%m-%d")
        baki_amt = total_bill - rec_amt
        remarks = st.text_input("Remarks", "Thank you for your business!")

        if st.button("💾 Save Bill & Export PDF", type="primary", use_container_width=True):
            if c_name and total_bill > 0 and s1:
                item_desc = s1 + (f" + {s2}" if s2 else "")
                DataManager.append_row("Invoices_Archive", {
                    "Invoice No": bill_no, "Date": bill_date, "Customer Name": c_name,
                    "Mobile Number": str(c_phone).strip(), "Service 1": s1, "Amount 1": amt1,
                    "Service 2": s2, "Amount 2": amt2, "Total Amount": total_bill,
                    "Paid Amount": rec_amt, "Pending Amount": baki_amt, "Payment Mode": pay_mode, "Remarks": remarks
                })
                if rec_amt > 0:
                    DataManager.append_row("Income", {"Date": bill_date, "Customer/Person": c_name, "Work Details": f"Bill #{bill_no}: {item_desc}", "Amount": rec_amt, "Payment Mode": pay_mode, "Notes": f"Mob: {c_phone}"})
                if baki_amt > 0:
                    DataManager.append_row("Udhar_Baki", {"Date": bill_date, "Customer Name": c_name, "Mobile Number": str(c_phone).strip(), "Service Details": item_desc, "Total Amount": total_bill, "Paid Amount": rec_amt, "Pending Amount": baki_amt, "Due Date": due_date, "Status": "Pending"})
                
                pdf_data = generate_invoice_pdf_buffer(bill_no, bill_date, c_name, c_phone, s1, amt1, s2, amt2, total_bill, rec_amt, baki_amt, pay_mode, remarks)
                col_dwn, col_wa = st.columns(2)
                col_dwn.download_button("📥 Download PDF Invoice", data=pdf_data, file_name=f"Invoice_{c_name}_{bill_no}.pdf", mime="application/pdf", type="primary", use_container_width=True)
                
                wa_msg = f"🧾 *TAX INVOICE*\n🏢 *{COMPANY_NAME}*\n📄 *Invoice No:* {bill_no}\n👤 *Customer:* {c_name}\n💼 *Service:* {item_desc}\n💰 *Total:* Rs. {total_bill:,.2f}\n✅ *Paid:* Rs. {rec_amt:,.2f}\n⚠️ *Pending:* Rs. {baki_amt:,.2f}\n📞 {COMPANY_MOBILE}"
                wa_url = f"https://wa.me/91{str(c_phone).strip()}?text={urllib.parse.quote(wa_msg)}"
                col_wa.markdown(f'<a href="{wa_url}" target="_blank"><button style="width:100%; height:45px; background-color:#25D366; color:white; border:none; border-radius:8px; font-weight:bold; cursor:pointer;">📲 Send Invoice via WhatsApp</button></a>', unsafe_allow_html=True)
                st.success("Invoice Saved Successfully!")

    # --- 2. EDIT / DELETE GENERATED INVOICES (PIN PROTECTED) ---
    elif bill_type == "✏️ Edit / Delete Invoices (Requires PIN)":
        st.markdown("### 🔐 Modify or Delete Existing Invoice Record")
        df_arch = DataManager.get_df("Invoices_Archive")
        
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
                        if inv_auth_pin == DataManager.get_pin():
                            DataManager.update_invoice_archive(sel_inv_no, {
                                "Date": up_date,
                                "Customer Name": up_cname,
                                "Mobile Number": str(up_cphone).strip(),
                                "Service 1": up_s1,
                                "Amount 1": up_amt1,
                                "Service 2": up_s2,
                                "Amount 2": up_amt2,
                                "Total Amount": up_tot,
                                "Paid Amount": up_rec,
                                "Pending Amount": up_baki,
                                "Payment Mode": up_mode,
                                "Remarks": up_remarks
                            })
                            st.success(f"✅ Invoice #{sel_inv_no} updated successfully!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN! Action denied.")
                            
                    if btn_del_col.button("🗑️ Delete Invoice Record", key=f"del_btn_{sel_inv_no}", type="primary", use_container_width=True):
                        if inv_auth_pin == DataManager.get_pin():
                            DataManager.delete_invoice_archive(sel_inv_no)
                            st.warning(f"🗑️ Invoice #{sel_inv_no} deleted successfully!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN! Action denied.")
        else:
            st.info("No generated invoice records found to edit.")

    # --- 3. RE-PRINT OLD INVOICES ---
    elif bill_type == "🖨️ Re-Print Old Invoice":
        df_arch = DataManager.get_df("Invoices_Archive")
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
        if st.button("💾 Save Expense", type="primary", use_container_width=True):
            if p_name and p_amt > 0:
                DataManager.append_row("Expense", {"Date": v_date, "Expense Name": f"{p_name} ({p_desc})", "Amount": p_amt, "Notes": f"VOU #{v_no} | {p_mode}"})
                st.success("Expense Recorded!")

    # --- 5. EDIT / DELETE EXPENSE VOUCHERS (PIN PROTECTED) ---
    elif bill_type == "✏️ Edit / Delete Vouchers (Requires PIN)":
        st.markdown("### 🔐 Modify or Delete Payment Voucher Record")
        df_exp = DataManager.get_df("Expense")
        
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
                        if exp_auth_pin == DataManager.get_pin():
                            DataManager.update_row("Expense", sel_exp_id, {
                                "Date": up_e_date,
                                "Expense Name": up_e_name,
                                "Amount": up_e_amt,
                                "Notes": up_e_notes
                            })
                            st.success(f"✅ Voucher #{sel_exp_id} updated successfully!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
                            
                    if eb_del_col.button("🗑️ Delete Voucher", key=f"del_exp_btn_{sel_exp_id}", type="primary", use_container_width=True):
                        if exp_auth_pin == DataManager.get_pin():
                            DataManager.delete_row("Expense", sel_exp_id)
                            st.warning(f"🗑️ Voucher #{sel_exp_id} deleted successfully!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
        else:
            st.info("No expense vouchers found to edit.")

    # --- 6. SETTLE OLD PENDING DUE ---
    elif bill_type == "Settle Old Pending Due":
        df_baki = DataManager.get_df("Udhar_Baki")
        if not df_baki.empty:
            pending = df_baki[df_baki["Pending Amount"] > 0]
            if not pending.empty:
                sel_acc = st.selectbox("Select Due Account:", [f"ID #{r['ID']} - {r['Customer Name']} | Pending: ₹{r['Pending Amount']}" for _, r in pending.iterrows()])
                sel_id = int(sel_acc.split(" ")[1].replace("#", ""))
                r = df_baki[df_baki["ID"] == sel_id].iloc[0]
                s_amt = st.number_input("Payment Received Now (₹) *", min_value=0.0, max_value=float(r["Pending Amount"]), value=float(r["Pending Amount"]), step=100.0)
                s_mode = st.selectbox("Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"])
                if st.button("💳 Settle Balance", type="primary", use_container_width=True):
                    new_paid = float(r["Paid Amount"]) + s_amt
                    new_pending = float(r["Pending Amount"]) - s_amt
                    DataManager.update_row("Udhar_Baki", sel_id, {"Paid Amount": new_paid, "Pending Amount": new_pending, "Status": "Cleared" if new_pending <= 0 else "Pending"})
                    DataManager.append_row("Income", {"Date": datetime.now().strftime("%Y-%m-%d"), "Customer/Person": r['Customer Name'], "Work Details": f"Due Settlement ({r.get('Service Details')})", "Amount": s_amt, "Payment Mode": s_mode, "Notes": f"Due Rec #{sel_id}"})
                    st.success("Due Settled!")
                    st.rerun()

# ----------------- 4. REPORTS & PDF -----------------
elif menu == "📄 Reports & PDF":
    st.subheader("📄 Financial Reports")
    c1, c2 = st.columns(2)
    d_from = c1.date_input("From Date", datetime.now().replace(day=1)).strftime("%Y-%m-%d")
    d_to = c2.date_input("To Date", datetime.now()).strftime("%Y-%m-%d")
    
    df_i = DataManager.get_df("Income")
    df_e = DataManager.get_df("Expense")
    df_b = DataManager.get_df("Udhar_Baki")
    
    f_i = df_i[(df_i["Date"] >= d_from) & (df_i["Date"] <= d_to)] if not df_i.empty and "Date" in df_i else pd.DataFrame()
    f_e = df_e[(df_e["Date"] >= d_from) & (df_e["Date"] <= d_to)] if not df_e.empty and "Date" in df_e else pd.DataFrame()
    
    t_i = f_i["Amount"].sum() if not f_i.empty and "Amount" in f_i else 0
    t_e = f_e["Amount"].sum() if not f_e.empty and "Amount" in f_e else 0
    cash_op, bank_op = DataManager.get_opening_balance()
    tot_op = cash_op + bank_op
    closing_bal = tot_op + t_i - t_e
    
    st.info(f"**Period:** {d_from} to {d_to} | **Revenue:** ₹{t_i:,.2f} | **Expenses:** ₹{t_e:,.2f} | **Closing Balance:** ₹{closing_bal:,.2f}")
    
    t1, t2, t3 = st.tabs(["💰 Income Report", "💸 Expense Report", "📋 Due Collections Report"])
    with t1:
        st.dataframe(f_i, use_container_width=True)
    with t2:
        st.dataframe(f_e, use_container_width=True)
    with t3:
        st.dataframe(df_b, use_container_width=True)

# ----------------- 5. OPENING BALANCE -----------------
elif menu == "🏦 Opening Balance":
    st.subheader("🏦 Opening Balance Setup")
    curr_c, curr_b = DataManager.get_opening_balance()
    c1, c2 = st.columns(2)
    in_c = c1.number_input("Cash in Hand (₹)", value=float(curr_c), step=500.0)
    in_b = c2.number_input("Bank Balance (₹)", value=float(curr_b), step=500.0)
    pin = st.text_input("Enter Security PIN to Save:", type="password")
    if st.button("💾 Save Opening Balance", type="primary", use_container_width=True):
        if pin == DataManager.get_pin():
            DataManager.set_opening_balance(in_c, in_b)
            st.success("Opening Balance Saved!")
            st.rerun()
        else:
            st.error("Invalid PIN!")

# ----------------- 6. CUSTOMERS DIRECTORY -----------------
elif menu == "👥 Customers Directory":
    st.subheader("👥 Client Directory & Broadcast")
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
                    df_c = DataManager.get_df("Customers")
                    clean_phone = str(cp).strip()
                    if not df_c.empty and "Mobile Number" in df_c and clean_phone in df_c["Mobile Number"].astype(str).values:
                        st.warning(f"⚠️ A customer with mobile {clean_phone} already exists in records!")
                    else:
                        DataManager.append_row("Customers", {
                            "Created Date": datetime.now().strftime("%Y-%m-%d"),
                            "Customer Name": cn,
                            "Mobile Number": clean_phone,
                            "City/Address": c_addr,
                            "Primary Service / Purpose": cs,
                            "Notes": c_notes
                        })
                        st.success(f"Client '{cn}' saved successfully!")
                        st.rerun()
                else:
                    st.error("Customer name and mobile number are required.")
                    
    with tab_list:
        df_c = DataManager.get_df("Customers")
        if not df_c.empty:
            search_query = st.text_input("🔍 Quick Search by Name, Mobile, Address or Service:", "")
            if search_query:
                filtered_df = df_c[
                    df_c["Customer Name"].astype(str).str.contains(search_query, case=False, na=False) | 
                    df_c["Mobile Number"].astype(str).str.contains(search_query, case=False, na=False) |
                    df_c.get("City/Address", pd.Series()).astype(str).str.contains(search_query, case=False, na=False) |
                    df_c.get("Primary Service / Purpose", pd.Series()).astype(str).str.contains(search_query, case=False, na=False)
                ]
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
                    u_caddr = ec3.text_input("Address / City / Village", str(c_row.get("City/Address", "Kadi")) if pd.notna(c_row.get("City/Address")) else "")
                    
                    curr_serv = str(c_row.get("Primary Service / Purpose", "VISA"))
                    serv_idx = SERVICE_OPTIONS.index(curr_serv) if curr_serv in SERVICE_OPTIONS else 0
                    u_cserv = ec4.selectbox("Primary Service", SERVICE_OPTIONS, index=serv_idx)
                    
                    u_cnotes = st.text_area("Notes / Remarks", str(c_row.get("Notes", "")) if pd.notna(c_row.get("Notes")) else "")
                    
                    st.markdown("🔒 **Security Confirmation:**")
                    edit_pin = st.text_input("Enter Master Security PIN:", type="password", key=f"c_pin_{sel_c_id}")
                    
                    b_col1, b_col2 = st.columns(2)
                    if b_col1.button("🔄 Update Customer Details", key=f"btn_up_{sel_c_id}", use_container_width=True):
                        if edit_pin == DataManager.get_pin():
                            DataManager.update_row("Customers", sel_c_id, {
                                "Customer Name": u_cname,
                                "Mobile Number": str(u_cphone).strip(),
                                "City/Address": u_caddr,
                                "Primary Service / Purpose": u_cserv,
                                "Notes": u_cnotes
                            })
                            st.success("Client profile updated successfully!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
                            
                    if b_col2.button("🗑️ Delete Customer", key=f"btn_del_{sel_c_id}", type="primary", use_container_width=True):
                        if edit_pin == DataManager.get_pin():
                            DataManager.delete_row("Customers", sel_c_id)
                            st.warning("Client deleted successfully!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
        else:
            st.info("No registered clients found.")

    with tab_promo:
        st.markdown("##### 📢 Bulk Broadcast & Promotion List")
        df_c = DataManager.get_df("Customers")
        if not df_c.empty:
            sel_aud = st.selectbox("Select Target Audience:", ["All Clients"] + list(df_c["Primary Service / Purpose"].dropna().unique()))
            target_df = df_c if sel_aud == "All Clients" else df_c[df_c["Primary Service / Purpose"] == sel_aud]
            
            st.write(f"**Total Recipients:** {len(target_df)}")
            st.dataframe(target_df[["Customer Name", "Mobile Number", "City/Address", "Primary Service / Purpose", "Notes"]], use_container_width=True)
            
            promo_msg = st.text_area("Broadcast Message Template:", value=f"Greetings from {COMPANY_NAME}! Contact us at {COMPANY_MOBILE} for special offers and updates regarding your service inquiry.")
            for _, prow in target_df.head(10).iterrows():
                p_url = f"https://wa.me/91{str(prow['Mobile Number']).strip()}?text={urllib.parse.quote(promo_msg)}"
                st.markdown(f"👉 **{prow['Customer Name']}** ({prow['Mobile Number']}) - [{prow.get('City/Address', 'Kadi')}]: [📲 Send WhatsApp]({p_url})")
        else:
            st.info("No client records available.")

# ----------------- 7. INCOME & EXPENSE MANAGEMENT -----------------
elif menu == "💰 Income":
    st.subheader("💰 Income Ledger")
    df_i = DataManager.get_df("Income")
    if not df_i.empty:
        st.dataframe(df_i, use_container_width=True)

elif menu == "💸 Expenses":
    st.subheader("💸 Expense Ledger")
    df_e = DataManager.get_df("Expense")
    if not df_e.empty:
        st.dataframe(df_e, use_container_width=True)

elif menu == "📋 Due Collections":
    st.subheader("📋 Due Collections Ledger")
    df_b = DataManager.get_df("Udhar_Baki")
    if not df_b.empty:
        st.dataframe(df_b, use_container_width=True)

# ----------------- 8. BACKUP & RESTORE DATA (PROTECTS DATA FROM CODE UPDATES) -----------------
elif menu == "💾 Backup & Restore":
    st.subheader("💾 Data Backup & Restore Center")
    st.info("💡 To prevent data loss during code updates, download a backup before updating, and upload it here after updating.")
    
    b_tab1, b_tab2 = st.tabs(["📥 Download Current Backup", "📤 Restore / Upload Previous Backup"])
    
    with b_tab1:
        st.markdown("##### 📥 Export All Accounting Data")
        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE, "rb") as f:
                st.download_button(
                    label="📥 Click to Download Full Database Backup (.xlsx)",
                    data=f,
                    file_name=f"Rojmed_Backup_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
        else:
            st.info("No data created yet.")
            
    with b_tab2:
        st.markdown("##### 📤 Restore Previous Database File")
        st.caption("Upload your previously downloaded `Rojmed_Backup.xlsx` or `Backup.xlsx` file here.")
        
        uploaded_backup = st.file_uploader("Choose Backup Excel File (.xlsx):", type=["xlsx"])
        rest_pin = st.text_input("Enter Master Security PIN to Confirm Restore:", type="password", key="rest_pin_inp")
        
        if st.button("🚀 Restore & Overwrite Data", type="primary", use_container_width=True):
            if rest_pin == DataManager.get_pin():
                if uploaded_backup is not None:
                    with open(EXCEL_FILE, "wb") as f:
                        f.write(uploaded_backup.getbuffer())
                    st.success("✅ Database successfully restored! All previous records are back.")
                    st.rerun()
                else:
                    st.error("Please select a valid backup excel file to upload.")
            else:
                st.error("❌ Incorrect Security PIN! Action denied.")

# ----------------- 9. SECURITY PIN -----------------
elif menu == "⚙️ Security / Change PIN":
    st.subheader("⚙️ Change Master PIN")
    with st.form("pin_form"):
        old_p = st.text_input("Current PIN *", type="password")
        new_p = st.text_input("New PIN *", type="password")
        conf_p = st.text_input("Confirm New PIN *", type="password")
        if st.form_submit_button("💾 Update PIN"):
            if old_p == DataManager.get_pin():
                if new_p and new_p == conf_p:
                    DataManager.set_pin(new_p)
                    st.success("PIN Updated!")
                else:
                    st.error("PIN mismatch!")
            else:
                st.error("Incorrect Current PIN!")
