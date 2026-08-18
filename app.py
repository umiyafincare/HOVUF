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
    ("🧾 Generate Bill / Voucher", "Create & Print Invoices"),
    ("📄 Reports & PDF", "Statements & Reports"),
    ("🏦 Opening Balance", "Set Starting Balances"),
    ("👥 Customers Directory", "Manage Clients & Broadcasts"),
    ("💰 Income", "View & Manage Income"),
    ("💸 Expenses", "View & Manage Expenses"),
    ("📋 Due Collections", "Customer Pending Dues"),
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

# Download Full Excel Backup
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button("📥 Backup All Data (Excel)", data=f, file_name=f"Backup_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

if st.sidebar.button("🔒 Logout System", use_container_width=True):
    st.session_state.logged_in = False
    st.rerun()

SERVICE_OPTIONS = [
    "VISA", "PASSPORT", "PANCARD", "LOAN", "LAND RECORD (7/12 & 8-A)",
    "E-KYC ALL", "FARMER REGISTRATION", "PM-KISAN", "DIGITAL GUJARAT SERVICES",
    "PROVIDENT FUND (PF)", "REVENUE WORK", "MARRIAGE CERTIFICATE", "INSURANCE",
    "AIR TICKET", "OTHER"
]

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

# ----------------- 3. INVOICE GENERATION -----------------
elif menu == "🧾 Generate Bill / Voucher":
    st.subheader("🧾 Generate & Re-Print Invoices")
    bill_type = st.radio("Select Action:", ["Customer Invoice (Income)", "🖨️ Re-Print Old Invoice", "Payment Voucher (Expense)", "Settle Old Pending Due"], horizontal=True)
    
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

    elif bill_type == "🖨️ Re-Print Old Invoice":
        df_arch = DataManager.get_df("Invoices_Archive")
        if not df_arch.empty:
            sel_inv = st.selectbox("Select Invoice:", [f"{r['Invoice No']} - {r['Customer Name']} ({r['Date']})" for _, r in df_arch.iterrows()])
            sel_no = sel_inv.split(" - ")[0]
            r = df_arch[df_arch["Invoice No"] == sel_no].iloc[0]
            re_pdf = generate_invoice_pdf_buffer(str(r["Invoice No"]), str(r["Date"]), str(r["Customer Name"]), str(r["Mobile Number"]), str(r.get("Service 1", "")), float(r.get("Amount 1", 0)), str(r.get("Service 2", "")), float(r.get("Amount 2", 0)), float(r["Total Amount"]), float(r.get("Paid Amount", 0)), float(r.get("Pending Amount", 0)), str(r.get("Payment Mode", "")), str(r.get("Remarks", "")))
            st.download_button("🖨️ Re-Download PDF", data=re_pdf, file_name=f"Invoice_{sel_no}.pdf", mime="application/pdf", type="primary", use_container_width=True)

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
    tab_new, tab_list = st.tabs(["➕ Add Client", "📋 Registered Clients"])
    with tab_new:
        with st.form("c_form"):
            cn = st.text_input("Customer Name *")
            cp = st.text_input("Mobile Number *")
            cs = st.selectbox("Primary Service", SERVICE_OPTIONS)
            if st.form_submit_button("💾 Save Client"):
                if cn and cp:
                    DataManager.append_row("Customers", {"Created Date": datetime.now().strftime("%Y-%m-%d"), "Customer Name": cn, "Mobile Number": cp, "Primary Service / Purpose": cs})
                    st.success("Client Saved!")
                    st.rerun()
    with tab_list:
        df_c = DataManager.get_df("Customers")
        if not df_c.empty:
            st.dataframe(df_c, use_container_width=True)

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

# ----------------- 8. SECURITY PIN -----------------
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
