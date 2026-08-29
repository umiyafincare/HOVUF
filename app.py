import os
import urllib.parse
from datetime import datetime, time
import io
import re
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image as PDFImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

COMPANY_NAME = "HARI OM VISA & UMIYA FINCARE"
COMPANY_ADDRESS = "F-46 VATSALY STATUS, NR. DHAVAL PLAZA, KADI - 384440"
COMPANY_MOBILE = "7698564672 / 9714776364"
COMPANY_TAGLINE = "Visa Consultancy | Insurance & Land Advisor | Property Solution | Daily Accounting"

# Official Payment Details
UPI_ID = "7698564672@upi"
PAYMENT_MOBILE = "7698564672"

# Image File Names
LOGO_VISA = "HARI OM.jpg"
LOGO_FINCARE = "UMIYA FIN.jpg"
LOGO_INSURANCE = "HARI OM IL.jpg"
LOGO_PROPERTY = "SHREE UNIYA.jpg"

DEFAULT_PIN = "1234"

st.set_page_config(page_title=COMPANY_NAME, page_icon="💼", layout="wide")

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

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    conn = None

DEFAULT_SCHEMAS = {
    "Customers": ["ID", "Created Date", "Customer Name", "Mobile Number", "City Address", "Primary Service", "Notes"],
    "Invoices_Archive": ["Invoice No", "Date", "Customer Name", "Mobile Number", "Service 1", "Amount 1", "Service 2", "Amount 2", "Total Amount", "Paid Amount", "Pending Amount", "Payment Mode", "Remarks"],
    "Income": ["ID", "Date", "Customer Person", "Work Details", "Amount", "Payment Mode", "Notes"],
    "Expense": ["ID", "Date", "Expense Name", "Amount", "Payment Mode", "Notes"],
    "Udhar_Baki": ["ID", "Date", "Customer Name", "Mobile Number", "Service Details", "Total Amount", "Paid Amount", "Pending Amount", "Due Date", "Status"],
    "Task_Reminder": ["ID", "Date", "Time", "Person Name", "Mobile", "Task Details", "Status"],
    "Settings": ["Setting Key", "Setting Value", "Updated Date"]
}

def clean_phone_number(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    digits = re.sub(r'\D', '', val_str)
    if len(digits) > 10 and digits.startswith("91"):
        digits = digits[2:]
    return digits

def format_to_ddmmyyyy(val):
    if pd.isna(val) or val is None or str(val).strip() == "" or str(val).strip() == "None":
        return ""
    val_str = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(val_str.split(" ")[0], fmt).strftime("%d/%m/%Y")
        except Exception:
            pass
    return val_str

def parse_date_safely(val_str):
    if not val_str or pd.isna(val_str):
        return datetime.now().date()
    val_clean = str(val_str).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(val_clean, fmt).date()
        except Exception:
            pass
    return datetime.now().date()

class GSheetsManager:
    @staticmethod
    def get_df(sheet_name):
        if conn is not None:
            try:
                df = conn.read(worksheet=sheet_name, ttl="30s")
                if df is not None and not df.empty:
                    df = df.dropna(how="all")
                    df.columns = [str(c).replace('_', ' ').strip() for c in df.columns]
                    
                    for c in df.columns:
                        if c != "ID" and not ("Amount" in c or "Balance" in c):
                            df[c] = df[c].astype(object)
                    
                    if "Mobile Number" in df.columns:
                        df["Mobile Number"] = df["Mobile Number"].apply(clean_phone_number)
                    if "Mobile" in df.columns:
                        df["Mobile"] = df["Mobile"].apply(clean_phone_number)
                    
                    for d_col in ["Date", "Created Date", "Due Date", "Updated Date"]:
                        if d_col in df.columns:
                            df[d_col] = df[d_col].apply(format_to_ddmmyyyy)
                    
                    if sheet_name == "Expense" and "Payment Mode" not in df.columns:
                        df["Payment Mode"] = "Cash"
                    
                    if sheet_name == "Customers":
                        if "Primary Service" not in df.columns:
                            for alt in ["Primary Service / Purpose", "Primary Service", "Service"]:
                                if alt in df.columns:
                                    df["Primary Service"] = df[alt]
                                    break
                        if "City Address" not in df.columns:
                            for alt in ["City/Address", "City Address", "Address", "City"]:
                                if alt in df.columns:
                                    df["City Address"] = df[alt]
                                    break
                    return df
            except Exception:
                pass
        cols = DEFAULT_SCHEMAS.get(sheet_name, ["ID"])
        return pd.DataFrame(columns=cols)

    @staticmethod
    def save_df(sheet_name, df):
        if conn is not None:
            try:
                conn.update(worksheet=sheet_name, data=df)
                st.cache_data.clear()
            except Exception as e:
                st.error(f"Google Sheets Sync Error: {e}")

    @staticmethod
    def append_row(sheet_name, row_dict):
        df = GSheetsManager.get_df(sheet_name)
        if "ID" in DEFAULT_SCHEMAS.get(sheet_name, []):
            new_id = 1 if df.empty or "ID" not in df.columns else (int(pd.to_numeric(df["ID"], errors='coerce').max()) + 1 if len(df["ID"].dropna()) > 0 else 1)
            row_dict["ID"] = new_id
        
        for d_col in ["Date", "Created Date", "Due Date", "Updated Date"]:
            if d_col in row_dict:
                row_dict[d_col] = format_to_ddmmyyyy(row_dict[d_col])
                
        if "Mobile Number" in row_dict:
            row_dict["Mobile Number"] = clean_phone_number(row_dict["Mobile Number"])
        if "Mobile" in row_dict:
            row_dict["Mobile"] = clean_phone_number(row_dict["Mobile"])
            
        clean_row = {k.replace('_', ' ').strip(): v for k, v in row_dict.items()}
        new_row_df = pd.DataFrame([clean_row])
        df = pd.concat([df, new_row_df], ignore_index=True)
        GSheetsManager.save_df(sheet_name, df)
        return row_dict.get("ID", 1)

    @staticmethod
    def update_row(sheet_name, row_id, updated_dict):
        df = GSheetsManager.get_df(sheet_name)
        if not df.empty and "ID" in df.columns:
            id_series = df["ID"].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            target_id = str(row_id).replace('.0', '').strip()
            idx = df.index[id_series == target_id].tolist()
            if idx:
                for k, v in updated_dict.items():
                    clean_k = k.replace('_', ' ').strip()
                    if clean_k in ["Date", "Created Date", "Due Date", "Updated Date"]:
                        v = format_to_ddmmyyyy(v)
                    if clean_k in ["Mobile Number", "Mobile"]:
                        v = clean_phone_number(v)
                    if clean_k not in df.columns:
                        df[clean_k] = None
                    if df[clean_k].dtype != object:
                        df[clean_k] = df[clean_k].astype(object)
                    df.at[idx[0], clean_k] = v
                GSheetsManager.save_df(sheet_name, df)

    @staticmethod
    def delete_row(sheet_name, row_id):
        df = GSheetsManager.get_df(sheet_name)
        if not df.empty and "ID" in df.columns:
            id_series = df["ID"].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            target_id = str(row_id).replace('.0', '').strip()
            df = df[id_series != target_id]
            GSheetsManager.save_df(sheet_name, df)

    @staticmethod
    def update_invoice(invoice_no, updated_dict):
        df = GSheetsManager.get_df("Invoices_Archive")
        if not df.empty and "Invoice No" in df.columns:
            idx = df.index[df["Invoice No"].astype(str).str.strip() == str(invoice_no).strip()].tolist()
            if idx:
                for k, v in updated_dict.items():
                    clean_k = k.replace('_', ' ').strip()
                    if clean_k == "Date":
                        v = format_to_ddmmyyyy(v)
                    if clean_k == "Mobile Number":
                        v = clean_phone_number(v)
                    if clean_k not in df.columns:
                        df[clean_k] = None
                    if df[clean_k].dtype != object:
                        df[clean_k] = df[clean_k].astype(object)
                    df.at[idx[0], clean_k] = v
                GSheetsManager.save_df("Invoices_Archive", df)

    @staticmethod
    def delete_invoice(invoice_no):
        df = GSheetsManager.get_df("Invoices_Archive")
        if not df.empty and "Invoice No" in df.columns:
            df = df[df["Invoice No"].astype(str).str.strip() != str(invoice_no).strip()]
            GSheetsManager.save_df("Invoices_Archive", df)

    @staticmethod
    def sync_customer(name, phone, service, remarks=""):
        clean_p = clean_phone_number(phone)
        if not clean_p or not name:
            return
        df_c = GSheetsManager.get_df("Customers")
        now_dt = datetime.now().strftime("%d/%m/%Y")
        if not df_c.empty and "Mobile Number" in df_c.columns:
            matched = df_c[df_c["Mobile Number"].astype(str).str.strip() == clean_p]
            if not matched.empty:
                c_id = matched.iloc[0]["ID"]
                GSheetsManager.update_row("Customers", c_id, {"Customer Name": name.strip(), "Primary Service": service})
                return
        GSheetsManager.append_row("Customers", {
            "Created Date": now_dt,
            "Customer Name": name.strip(),
            "Mobile Number": clean_p,
            "City Address": "Kadi",
            "Primary Service": service,
            "Notes": remarks
        })

    @staticmethod
    def get_pin():
        try:
            df_s = GSheetsManager.get_df("Settings")
            if not df_s.empty:
                key_col, val_col = None, None
                for col in df_s.columns:
                    c_clean = str(col).lower().replace(" ", "").replace("_", "")
                    if "key" in c_clean:
                        key_col = col
                    if "value" in c_clean:
                        val_col = col
                if key_col and val_col:
                    row = df_s[df_s[key_col].astype(str).str.strip() == "Master_PIN"]
                    if not row.empty:
                        raw_val = row.iloc[0][val_col]
                        if pd.notna(raw_val):
                            val_str = str(raw_val).strip()
                            if val_str.endswith(".0"):
                                val_str = val_str[:-2]
                            return str(val_str)
        except Exception:
            pass
        return DEFAULT_PIN

    @staticmethod
    def set_pin(new_pin):
        df_s = GSheetsManager.get_df("Settings")
        now_str = datetime.now().strftime("%d/%m/%Y")
        clean_p = str(new_pin).strip()
        if df_s.empty or "Setting Key" not in df_s.columns:
            df_s = pd.DataFrame([{"Setting Key": "Master_PIN", "Setting Value": clean_p, "Updated Date": now_str}])
        else:
            if "Master_PIN" in df_s["Setting Key"].values:
                df_s.loc[df_s["Setting Key"] == "Master_PIN", "Setting Value"] = clean_p
                df_s.loc[df_s["Setting Key"] == "Master_PIN", "Updated Date"] = now_str
            else:
                df_s = pd.concat([df_s, pd.DataFrame([{"Setting Key": "Master_PIN", "Setting Value": clean_p, "Updated Date": now_str}])], ignore_index=True)
        GSheetsManager.save_df("Settings", df_s)

    @staticmethod
    def get_opening_balance():
        df_s = GSheetsManager.get_df("Settings")
        cash_op, bank_op = 0.0, 0.0
        if not df_s.empty and "Setting Key" in df_s.columns and "Setting Value" in df_s.columns:
            rc = df_s[df_s["Setting Key"] == "Cash_Opening_Balance"]
            rb = df_s[df_s["Setting Key"] == "Bank_Opening_Balance"]
            if not rc.empty and pd.notna(rc.iloc[0]["Setting Value"]):
                try:
                    cash_op = float(rc.iloc[0]["Setting Value"])
                except Exception:
                    pass
            if not rb.empty and pd.notna(rb.iloc[0]["Setting Value"]):
                try:
                    bank_op = float(rb.iloc[0]["Setting Value"])
                except Exception:
                    pass
        return cash_op, bank_op

    @staticmethod
    def set_opening_balance(cash_op, bank_op):
        df_s = GSheetsManager.get_df("Settings")
        now_str = datetime.now().strftime("%d/%m/%Y")
        recs = [
            {"Setting Key": "Cash_Opening_Balance", "Setting Value": str(cash_op), "Updated Date": now_str},
            {"Setting Key": "Bank_Opening_Balance", "Setting Value": str(bank_op), "Updated Date": now_str}
        ]
        if df_s.empty or "Setting Key" not in df_s.columns:
            df_s = pd.DataFrame(recs)
        else:
            for r in recs:
                if r["Setting Key"] in df_s["Setting Key"].values:
                    df_s.loc[df_s["Setting Key"] == r["Setting Key"], "Setting Value"] = r["Setting Value"]
                    df_s.loc[df_s["Setting Key"] == r["Setting Key"], "Updated Date"] = now_str
                else:
                    df_s = pd.concat([df_s, pd.DataFrame([r])], ignore_index=True)
        GSheetsManager.save_df("Settings", df_s)

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

# ----------------- 🔒 SECURE LOGIN SCREEN -----------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    render_top_logos()
    col_c1, col_c2, col_c3 = st.columns([1, 1.2, 1])
    with col_c2:
        st.markdown("""
            <div style="background: #FFFFFF; padding: 25px; border-radius: 12px; border: 1px solid #E2E8F0; text-align: center;">
                <h3 style="color: #1E3A8A; margin-top: 0;">🔒 Secure Ledger Login</h3>
                <p style="color: #64748B; font-size: 13px;">Enter your Security Master PIN to access the accounting system.</p>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        entered_pin = st.text_input("Enter Security PIN:", type="password", max_chars=12)
        
        if st.button("🔓 Unlock & Login", type="primary", use_container_width=True):
            saved_pin = str(GSheetsManager.get_pin()).strip()
            user_input = str(entered_pin).strip()
            
            if user_input != "" and user_input == saved_pin:
                st.session_state.logged_in = True
                st.success("Access Granted!")
                st.rerun()
            else:
                st.error("❌ Invalid PIN! Please enter the correct Security PIN.")
    st.stop()

render_top_logos()

# ----------------- NAVIGATION MENU -----------------
if "current_page" not in st.session_state:
    st.session_state.current_page = "📊 Dashboard"

st.sidebar.markdown("<h4 style='color:#1E3A8A;'>📌 Navigation Menu</h4>", unsafe_allow_html=True)
menu_items = [
    ("📊 Dashboard", "Business metrics & live cash/bank summary"),
    ("⏰ Task Reminders", "Calendar & Clock Reminders"),
    ("🧾 Generate Bill / Voucher", "Create, Edit, Print Invoices & Vouchers"),
    ("📋 Due Collections", "Customer Pending Dues & Edit Records"),
    ("📄 Reports & PDF", "Financial Statements & PDF Export"),
    ("🏦 Opening Balance", "Set Starting Balances"),
    ("👥 Customers Directory", "Manage Clients & Broadcasts"),
    ("💰 Income", "View & Manage Income"),
    ("💸 Expenses", "View & Manage Expenses"),
    ("💾 Cloud Excel Backup", "Download Complete Data"),
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

if st.sidebar.button("🔄 Refresh Cloud Data", use_container_width=True):
    st.cache_data.clear()
    st.success("Data Refreshed!")
    st.rerun()

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
    clean_q = clean_phone_number(search_text) if any(c.isdigit() for c in str(search_text)) else str(search_text).strip().lower()
    df_c = GSheetsManager.get_df("Customers")
    df_b = GSheetsManager.get_df("Udhar_Baki")
    
    matched_cust = None
    if not df_c.empty and "Mobile Number" in df_c.columns:
        m_phone = df_c[df_c["Mobile Number"].astype(str).str.strip() == clean_q]
        if not m_phone.empty:
            matched_cust = m_phone.iloc[0]
        elif "Customer Name" in df_c.columns:
            m_name = df_c[df_c["Customer Name"].astype(str).str.lower() == clean_q]
            if not m_name.empty:
                matched_cust = m_name.iloc[0]
                
    total_due = 0.0
    due_records = []
    if not df_b.empty and "Pending Amount" in df_b.columns:
        cond = (df_b["Mobile Number"].astype(str).str.strip() == clean_q) if "Mobile Number" in df_b.columns else pd.Series([False]*len(df_b))
        if "Customer Name" in df_b.columns:
            cond = cond | (df_b["Customer Name"].astype(str).str.lower() == clean_q)
        m_due = df_b[cond]
        active_dues = m_due[pd.to_numeric(m_due["Pending Amount"], errors='coerce') > 0]
        if not active_dues.empty:
            total_due = float(pd.to_numeric(active_dues["Pending Amount"], errors='coerce').sum())
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
    
    clean_p_disp = clean_phone_number(cust_phone)
    status_text = "PAID" if baki_amt <= 0 else f"PARTIAL (Balance: Rs. {baki_amt:,.2f})"
    meta_data = [
        [f"<b>Invoice No:</b> {bill_no}", f"<b>Date:</b> {format_to_ddmmyyyy(bill_date)}"],
        [f"<b>Name:</b> {cust_name}", f"<b>Mobile:</b> {clean_p_disp}"],
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
    elems.append(Spacer(1, 12))
    
    # Online Payment Box in PDF Bill (If balance due)
    if baki_amt > 0:
        upi_info = Paragraph(f"<b>Online Payment Details (GPay / PhonePe / Paytm / BHIM):</b><br/>Mobile No: <b>{PAYMENT_MOBILE}</b> | UPI ID: <b>{UPI_ID}</b><br/>Pending Balance: <b>Rs. {baki_amt:,.2f}</b>", styles['Normal'])
        t_upi = Table([[upi_info]], colWidths=[550])
        t_upi.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F1F5F9")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")), ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6)]))
        elems.append(t_upi)
        elems.append(Spacer(1, 10))
    
    foot_data = [[f"<b>Note:</b> {remarks}", f"For, <b>{COMPANY_NAME}</b>"], ["", "\n\n___________________________\nAuthorized Signature"]]
    t_foot = Table([[Paragraph(c, styles['Normal']) for c in r] for r in foot_data], colWidths=[310, 240])
    t_foot.setStyle(TableStyle([('ALIGN', (1, 0), (1, -1), 'RIGHT'), ('VALIGN', (0, 0), (-1, -1), 'TOP')]))
    elems.append(t_foot)
    doc.build(elems)
    buf.seek(0)
    return buf

# ----------------- 1. DASHBOARD (LIVE CASH & BANK SEPARATION) -----------------
if menu == "📊 Dashboard":
    st.subheader("📊 Business Overview & Live Balances (Google Sheets)")
    
    # Active Reminders
    df_rem_all = GSheetsManager.get_df("Task_Reminder")
    if not df_rem_all.empty and "Status" in df_rem_all.columns:
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
                col_t1.write(f"📅 **{format_to_ddmmyyyy(tr.get('Date'))}**")
                col_t2.write(f"⏰ {tr.get('Time')}")
                col_t3.write(f"📌 **{tr.get('Task Details')}** ({tr.get('Person Name', 'Client')})")
                clean_rem_mob = clean_phone_number(tr.get('Mobile'))
                if clean_rem_mob and len(clean_rem_mob) >= 10:
                    t_msg = f"Reminder regarding: {tr.get('Task Details')} scheduled on {format_to_ddmmyyyy(tr.get('Date'))} at {tr.get('Time')}."
                    t_url = f"https://wa.me/91{clean_rem_mob}?text={urllib.parse.quote(t_msg)}"
                    col_t4.markdown(f"[📲 WhatsApp]({t_url})")
                else:
                    col_t4.write("-")
                if col_t5.button("✅ Complete", key=f"dash_comp_{t_id}"):
                    GSheetsManager.update_row("Task_Reminder", t_id, {"Status": "Completed"})
                    st.success("Task Completed!")
                    st.rerun()
            st.divider()

    df_inc = GSheetsManager.get_df("Income")
    df_exp = GSheetsManager.get_df("Expense")
    df_baki = GSheetsManager.get_df("Udhar_Baki")
    df_cust = GSheetsManager.get_df("Customers")
    today_str = datetime.now().strftime("%d/%m/%Y")
    
    # Mode-wise Income Separation
    cash_inc_total = 0.0
    bank_inc_total = 0.0
    today_inc_total = 0.0
    
    if not df_inc.empty and "Amount" in df_inc.columns:
        df_inc["Amount_Num"] = pd.to_numeric(df_inc["Amount"], errors='coerce').fillna(0.0)
        mode_col = df_inc["Payment Mode"].astype(str).str.lower() if "Payment Mode" in df_inc.columns else pd.Series(["cash"]*len(df_inc))
        
        cash_inc_total = df_inc[mode_col.str.contains("cash", na=False)]["Amount_Num"].sum()
        bank_inc_total = df_inc[~mode_col.str.contains("cash", na=False)]["Amount_Num"].sum()
        
        if "Date" in df_inc.columns:
            today_inc_total = df_inc[df_inc["Date"] == today_str]["Amount_Num"].sum()

    # Mode-wise Expense Separation
    cash_exp_total = 0.0
    bank_exp_total = 0.0
    today_exp_total = 0.0
    
    if not df_exp.empty and "Amount" in df_exp.columns:
        df_exp["Amount_Num"] = pd.to_numeric(df_exp["Amount"], errors='coerce').fillna(0.0)
        
        if "Payment Mode" in df_exp.columns:
            exp_mode_col = df_exp["Payment Mode"].astype(str).str.lower()
        elif "Notes" in df_exp.columns:
            exp_mode_col = df_exp["Notes"].astype(str).str.lower()
        else:
            exp_mode_col = pd.Series(["cash"]*len(df_exp))
            
        cash_exp_total = df_exp[exp_mode_col.str.contains("cash", na=False)]["Amount_Num"].sum()
        bank_exp_total = df_exp[~exp_mode_col.str.contains("cash", na=False)]["Amount_Num"].sum()
        
        if "Date" in df_exp.columns:
            today_exp_total = df_exp[df_exp["Date"] == today_str]["Amount_Num"].sum()

    total_inc = cash_inc_total + bank_inc_total
    total_exp = cash_exp_total + bank_exp_total
    total_baki = pd.to_numeric(df_baki["Pending Amount"], errors='coerce').sum() if not df_baki.empty and "Pending Amount" in df_baki.columns else 0.0
    
    cash_op, bank_op = GSheetsManager.get_opening_balance()
    
    current_cash_in_hand = cash_op + cash_inc_total - cash_exp_total
    current_bank_balance = bank_op + bank_inc_total - bank_exp_total
    closing_net_balance = current_cash_in_hand + current_bank_balance
    total_cust = len(df_cust) if not df_cust.empty else 0

    st.markdown("#### 💼 Real-Time Available Funds")
    r1_c1, r1_c2, r1_c3 = st.columns(3)
    with r1_c1:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #16A34A;">
                <div class="kpi-label">💵 Live Cash in Hand (ગલ્લો)</div>
                <div class="kpi-value" style="color: #15803D;">₹ {current_cash_in_hand:,.2f}</div>
                <div class="kpi-sub">Opening: ₹ {cash_op:,.0f} | In: +₹ {cash_inc_total:,.0f} | Out: -₹ {cash_exp_total:,.0f}</div>
            </div>
        """, unsafe_allow_html=True)
    with r1_c2:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #2563EB;">
                <div class="kpi-label">🏦 Live Bank / UPI Balance</div>
                <div class="kpi-value" style="color: #1D4ED8;">₹ {current_bank_balance:,.2f}</div>
                <div class="kpi-sub">Opening: ₹ {bank_op:,.0f} | In: +₹ {bank_inc_total:,.0f} | Out: -₹ {bank_exp_total:,.0f}</div>
            </div>
        """, unsafe_allow_html=True)
    with r1_c3:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #0284C7;">
                <div class="kpi-label">💼 Total Net Balance (કુલ સિલક)</div>
                <div class="kpi-value" style="color: #0369A1;">₹ {closing_net_balance:,.2f}</div>
                <div class="kpi-sub">Cash + Bank Combined Capital</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("#### 📈 Revenue & Activity Summary")
    r2_c1, r2_c2, r2_c3 = st.columns(3)
    with r2_c1:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #059669;">
                <div class="kpi-label">💰 Total Revenue (Income)</div>
                <div class="kpi-value" style="color: #047857;">₹ {total_inc:,.2f}</div>
                <div class="kpi-sub">Cash: ₹ {cash_inc_total:,.0f} | Online: ₹ {bank_inc_total:,.0f} (Today: ₹ {today_inc_total:,.0f})</div>
            </div>
        """, unsafe_allow_html=True)
    with r2_c2:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #DC2626;">
                <div class="kpi-label">💸 Total Expenses (ખર્ચ)</div>
                <div class="kpi-value" style="color: #DC2626;">₹ {total_exp:,.2f}</div>
                <div class="kpi-sub">Cash: ₹ {cash_exp_total:,.0f} | Online: ₹ {bank_exp_total:,.0f} (Today: ₹ {today_exp_total:,.0f})</div>
            </div>
        """, unsafe_allow_html=True)
    with r2_c3:
        st.markdown(f"""
            <div class="kpi-card" style="border-left: 5px solid #D97706;">
                <div class="kpi-label">📋 Total Pending Dues</div>
                <div class="kpi-value" style="color: #B45309;">₹ {total_baki:,.2f}</div>
                <div class="kpi-sub">Clients Count: {total_cust} registered</div>
            </div>
        """, unsafe_allow_html=True)

    st.divider()
    st.subheader("📋 Pending Collections & WhatsApp Reminders")

    if not df_baki.empty and "Pending Amount" in df_baki.columns:
        pending = df_baki[pd.to_numeric(df_baki["Pending Amount"], errors='coerce') > 0]
        if not pending.empty:
            for _, r in pending.iterrows():
                b1, b2, b3, b4, b5, b6 = st.columns([2, 2, 2, 2, 2, 2])
                b1.write(f"**{r.get('Customer Name')}**")
                
                clean_phone = clean_phone_number(r.get('Mobile Number'))
                b2.write(f"📞 {clean_phone}")
                
                serv_name = str(r.get('Service Details', 'Service'))
                b3.write(f"🏷️ *{serv_name}*")
                due_val = float(pd.to_numeric(r.get('Pending Amount', 0), errors='coerce'))
                b4.write(f"Due: **₹ {due_val:,.2f}**")
                b5.write(f"Date: {format_to_ddmmyyyy(r.get('Due Date'))}")
                
                # Professional WhatsApp Message with Mobile No & UPI ID for Payment
                msg = (
                    f"🙏 *નમસ્તે {r.get('Customer Name')}*,\n\n"
                    f"🏢 *{COMPANY_NAME}*\n"
                    f"📌 *વિગત:* {serv_name}\n"
                    f"💰 *બાકી રકમ (Pending Due):* ₹ {due_val:,.2f}\n"
                    f"📅 *તારીખ:* {format_to_ddmmyyyy(r.get('Due Date'))}\n\n"
                    f"💳 *GPay / PhonePe / BHIM UPI પેમેન્ટ માટે:*\n"
                    f"👉 *મોબાઈલ નંબર:* `{PAYMENT_MOBILE}`\n"
                    f"👉 *UPI ID:* `{UPI_ID}`\n\n"
                    f"📞 *સંપર્ક:* {COMPANY_MOBILE}\n"
                    f"આભાર!"
                )
                
                if clean_phone and len(clean_phone) >= 10:
                    wa_url = f"https://wa.me/91{clean_phone}?text={urllib.parse.quote(msg)}"
                    b6.markdown(f"[📲 Send WhatsApp Reminder]({wa_url})", unsafe_allow_html=True)
                else:
                    b6.write("⚠️ Invalid Number")
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
            rdate = c4.date_input("📅 Reminder Date", datetime.now(), format="DD/MM/YYYY").strftime("%d/%m/%Y")
            rtime = st.time_input("⏰ Reminder Time", time(11, 0)).strftime("%I:%M %p")
            if st.form_submit_button("💾 Save Reminder", use_container_width=True):
                if tdesc and pname:
                    GSheetsManager.append_row("Task_Reminder", {
                        "Date": rdate, "Time": rtime, "Person Name": pname,
                        "Mobile": clean_phone_number(rphone), "Task Details": tdesc, "Status": "Pending"
                    })
                    st.success("Reminder Saved to Google Sheets!")
                    st.rerun()

    with tab_pending_tasks:
        df_rem = GSheetsManager.get_df("Task_Reminder")
        if not df_rem.empty and "Status" in df_rem.columns:
            pending_list = df_rem[df_rem["Status"] == "Pending"]
            for _, r in pending_list.iterrows():
                r_id = r["ID"]
                c1, c2, c3, c4 = st.columns([2, 2, 3, 2])
                c1.write(f"📅 {format_to_ddmmyyyy(r['Date'])} | ⏰ {r['Time']}")
                c2.write(f"👤 {r.get('Person Name')}")
                c3.write(f"📌 {r.get('Task Details')}")
                clean_rem_mob = clean_phone_number(r.get('Mobile'))
                if clean_rem_mob and len(clean_rem_mob) >= 10:
                    t_msg = f"Reminder regarding: {r.get('Task Details')} scheduled on {format_to_ddmmyyyy(r.get('Date'))} at {r.get('Time')}."
                    t_url = f"https://wa.me/91{clean_rem_mob}?text={urllib.parse.quote(t_msg)}"
                    c4.markdown(f"[📲 WhatsApp]({t_url})")
                if c4.button("✅ Done", key=f"done_{r_id}"):
                    GSheetsManager.update_row("Task_Reminder", r_id, {"Status": "Completed"})
                    st.rerun()

    with tab_completed_tasks:
        df_rem = GSheetsManager.get_df("Task_Reminder")
        if not df_rem.empty and "Status" in df_rem.columns:
            st.dataframe(df_rem[df_rem["Status"] == "Completed"], use_container_width=True)

# ----------------- 3. INVOICE GENERATION -----------------
elif menu == "🧾 Generate Bill / Voucher":
    st.subheader("🧾 Generate, Edit & Manage Invoices / Vouchers (Cloud)")
    bill_type = st.radio("Select Action:", [
        "Customer Invoice (Income)", 
        "✏️ Edit / Delete Invoices (Requires PIN)",
        "🖨️ Re-Print Old Invoice", 
        "Payment Voucher (Expense)", 
        "✏️ Edit / Delete Vouchers (Requires PIN)",
        "Settle Old Pending Due"
    ], horizontal=True)
    
    if bill_type == "Customer Invoice (Income)":
        st.markdown("### 🔍 STEP 1: Quick Customer Lookup & Live Due Detection")
        
        df_all_cust = GSheetsManager.get_df("Customers")
        col_s_opt1, col_s_opt2 = st.columns([1.5, 2.5])
        
        selected_from_list = None
        if not df_all_cust.empty:
            serv_col = "Primary Service" if "Primary Service" in df_all_cust.columns else df_all_cust.columns[min(5, len(df_all_cust.columns)-1)]
            cust_quick_list = ["-- Quick Choose Registered Client (Optional) --"] + [
                f"{r.get('Customer Name', '')} ({clean_phone_number(r.get('Mobile Number', ''))}) - {r.get(serv_col, '')}" 
                for _, r in df_all_cust.iterrows()
            ]
            chosen_c = col_s_opt1.selectbox("Search from Directory:", cust_quick_list)
            if chosen_c != "-- Quick Choose Registered Client (Optional) --":
                c_idx = cust_quick_list.index(chosen_c) - 1
                selected_from_list = df_all_cust.iloc[c_idx]

        init_name = str(selected_from_list.get("Customer Name", "")) if selected_from_list is not None else ""
        init_phone = clean_phone_number(selected_from_list.get("Mobile Number", "")) if selected_from_list is not None else ""
        init_service = str(selected_from_list.get("Primary Service", "VISA")) if selected_from_list is not None else "VISA"
        
        col_c1, col_c2 = st.columns(2)
        cust_name = col_c1.text_input("Customer Name *", value=init_name)
        cust_phone = col_c2.text_input("Mobile Number (10 Digits) *", value=init_phone)
        
        search_term = cust_phone if cust_phone else cust_name
        if search_term:
            matched_profile, live_due, due_records = search_customer_profile(search_term)
            
            if matched_profile is not None or live_due > 0:
                p_name = matched_profile.get('Customer Name') if matched_profile is not None else cust_name
                p_phone = clean_phone_number(matched_profile.get('Mobile Number')) if matched_profile is not None else clean_phone_number(cust_phone)
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
        bill_date = c4.date_input("Invoice Date", datetime.now(), format="DD/MM/YYYY").strftime("%d/%m/%Y")
        
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
        due_date = cp3.date_input("Due Date (If balance pending)", datetime.now(), format="DD/MM/YYYY").strftime("%d/%m/%Y")
        baki_amt = total_bill - rec_amt
        item_desc = s1 + (f" + {s2}" if s2 else "")
        
        if baki_amt > 0:
            st.warning(f"⚠️ Balance Due on this Bill: ₹ {baki_amt:,.2f} for '{item_desc}'")
            
        remarks = st.text_input("Remarks / Notes", "Thank you for choosing our services!")

        if st.button("💾 Generate Bill & Save to Cloud", type="primary", use_container_width=True):
            if cust_name and total_bill > 0 and s1:
                clean_phone_val = clean_phone_number(cust_phone)
                GSheetsManager.sync_customer(cust_name, clean_phone_val, s1, remarks)
                
                GSheetsManager.append_row("Invoices_Archive", {
                    "Invoice No": bill_no, "Date": bill_date, "Customer Name": cust_name.strip(),
                    "Mobile Number": clean_phone_val, "Service 1": s1, "Amount 1": amt1,
                    "Service 2": s2, "Amount 2": amt2, "Total Amount": total_bill,
                    "Paid Amount": rec_amt, "Pending Amount": baki_amt, "Payment Mode": pay_mode, "Remarks": remarks
                })
                
                if rec_amt > 0:
                    GSheetsManager.append_row("Income", {
                        "Date": bill_date, "Customer Person": cust_name.strip(),
                        "Work Details": f"Bill #{bill_no}: {item_desc}", "Amount": rec_amt,
                        "Payment Mode": pay_mode, "Notes": f"Mob: {clean_phone_val}"
                    })
                
                if baki_amt > 0:
                    GSheetsManager.append_row("Udhar_Baki", {
                        "Date": bill_date, "Customer Name": cust_name.strip(),
                        "Mobile Number": clean_phone_val, "Service Details": item_desc,
                        "Total Amount": total_bill, "Paid Amount": rec_amt,
                        "Pending Amount": baki_amt, "Due Date": due_date, "Status": "Pending"
                    })
                
                pdf_data = generate_invoice_pdf_buffer(bill_no, bill_date, cust_name, clean_phone_val, s1, amt1, s2, amt2, total_bill, rec_amt, baki_amt, pay_mode, remarks)
                col_dwn, col_wa = st.columns(2)
                col_dwn.download_button("📥 Download PDF Invoice", data=pdf_data, file_name=f"Invoice_{cust_name}_{bill_no}.pdf", mime="application/pdf", type="primary", use_container_width=True)
                
                wa_msg = (
                    f"🧾 *TAX INVOICE / RECEIPT*\n"
                    f"🏢 *{COMPANY_NAME}*\n"
                    f"📄 *Invoice No:* {bill_no}\n"
                    f"📅 *Date:* {bill_date}\n"
                    f"👤 *Customer:* {cust_name}\n"
                    f"💼 *Service:* {item_desc}\n"
                    f"💰 *Total:* Rs. {total_bill:,.2f}\n"
                    f"✅ *Paid:* Rs. {rec_amt:,.2f} ({pay_mode})\n"
                )
                if baki_amt > 0:
                    wa_msg += (
                        f"⚠️ *Pending Due:* Rs. {baki_amt:,.2f} (Due: {due_date})\n\n"
                        f"💳 *GPay / PhonePe / BHIM UPI પેમેન્ટ માટે:*\n"
                        f"👉 *મોબાઈલ નંબર:* `{PAYMENT_MOBILE}`\n"
                        f"👉 *UPI ID:* `{UPI_ID}`\n\n"
                    )
                wa_msg += f"📞 {COMPANY_MOBILE}\n🙏 *Thank you for your business!*"
                
                if clean_phone_val and len(clean_phone_val) >= 10:
                    wa_url = f"https://wa.me/91{clean_phone_val}?text={urllib.parse.quote(wa_msg)}"
                    col_wa.markdown(f'<a href="{wa_url}" target="_blank"><button style="width:100%; height:45px; background-color:#25D366; color:white; border:none; border-radius:8px; font-weight:bold; cursor:pointer;">📲 Send Invoice via WhatsApp</button></a>', unsafe_allow_html=True)
                st.success("✅ Bill Created & Saved to Google Sheets!")
            else:
                st.error("Please enter customer name, valid service, and bill amount.")

    elif bill_type == "✏️ Edit / Delete Invoices (Requires PIN)":
        st.markdown("### 🔐 Modify or Delete Existing Invoice Record (Cloud)")
        df_arch = GSheetsManager.get_df("Invoices_Archive")
        
        if not df_arch.empty and "Invoice No" in df_arch.columns:
            sel_inv_options = [f"{r['Invoice No']} - {r['Customer Name']} ({format_to_ddmmyyyy(r['Date'])}) | Total: ₹{r['Total Amount']}" for _, r in df_arch.iterrows()]
            chosen_inv_str = st.selectbox("Select Invoice to Modify / Delete:", sel_inv_options, key="edit_inv_select")
            
            if chosen_inv_str:
                sel_inv_no = chosen_inv_str.split(" - ")[0]
                inv_row = df_arch[df_arch["Invoice No"] == sel_inv_no].iloc[0]
                
                with st.expander(f"📝 Edit Invoice #{sel_inv_no} Details", expanded=True):
                    ed_c1, ed_c2 = st.columns(2)
                    raw_inv_dt = parse_date_safely(inv_row.get("Date", ""))
                    up_date = ed_c1.date_input("Invoice Date", raw_inv_dt, format="DD/MM/YYYY", key=f"inv_dt_inp_{sel_inv_no}").strftime("%d/%m/%Y")
                    up_cname = ed_c2.text_input("Customer Name", str(inv_row.get("Customer Name", "")))
                    
                    ed_c3, ed_c4 = st.columns(2)
                    up_cphone = ed_c3.text_input("Mobile Number", clean_phone_number(inv_row.get("Mobile Number", "")))
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
                    inv_auth_pin = st.text_input("Enter Security PIN to Authorize:", type="password", key=f"inv_pin_{sel_inv_no}")
                    
                    btn_up_col, btn_del_col = st.columns(2)
                    if btn_up_col.button("🔄 Update Invoice Record", key=f"up_btn_{sel_inv_no}", use_container_width=True):
                        if inv_auth_pin == GSheetsManager.get_pin():
                            GSheetsManager.update_invoice(sel_inv_no, {
                                "Date": str(up_date),
                                "Customer Name": str(up_cname),
                                "Mobile Number": clean_phone_number(up_cphone),
                                "Service 1": str(up_s1),
                                "Amount 1": float(up_amt1),
                                "Service 2": str(up_s2),
                                "Amount 2": float(up_amt2),
                                "Total Amount": float(up_tot),
                                "Paid Amount": float(up_rec),
                                "Pending Amount": float(up_baki),
                                "Payment Mode": str(up_mode),
                                "Remarks": str(up_remarks)
                            })
                            st.success(f"✅ Invoice #{sel_inv_no} updated in Google Sheets!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
                            
                    if btn_del_col.button("🗑️ Delete Invoice Record", key=f"del_btn_{sel_inv_no}", type="primary", use_container_width=True):
                        if inv_auth_pin == GSheetsManager.get_pin():
                            GSheetsManager.delete_invoice(sel_inv_no)
                            st.warning(f"🗑️ Invoice #{sel_inv_no} deleted from Google Sheets!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
        else:
            st.info("No generated invoice records found.")

    elif bill_type == "🖨️ Re-Print Old Invoice":
        df_arch = GSheetsManager.get_df("Invoices_Archive")
        if not df_arch.empty and "Invoice No" in df_arch.columns:
            sel_inv = st.selectbox("Select Invoice:", [f"{r['Invoice No']} - {r['Customer Name']} ({format_to_ddmmyyyy(r['Date'])})" for _, r in df_arch.iterrows()])
            sel_no = sel_inv.split(" - ")[0]
            r = df_arch[df_arch["Invoice No"] == sel_no].iloc[0]
            re_pdf = generate_invoice_pdf_buffer(str(r["Invoice No"]), str(r["Date"]), str(r["Customer Name"]), clean_phone_number(r["Mobile Number"]), str(r.get("Service 1", "")), float(r.get("Amount 1", 0)), str(r.get("Service 2", "")), float(r.get("Amount 2", 0)), float(r["Total Amount"]), float(r.get("Paid Amount", 0)), float(r.get("Pending Amount", 0)), str(r.get("Payment Mode", "")), str(r.get("Remarks", "")))
            st.download_button("🖨️ Re-Download PDF", data=re_pdf, file_name=f"Invoice_{sel_no}.pdf", mime="application/pdf", type="primary", use_container_width=True)
        else:
            st.info("No invoices found.")

    elif bill_type == "Payment Voucher (Expense)":
        c1, c2 = st.columns(2)
        v_no = c1.text_input("Voucher No.", f"VOU-{datetime.now().strftime('%Y%m%d%H%M')}")
        v_date = c2.date_input("Date", datetime.now(), format="DD/MM/YYYY").strftime("%d/%m/%Y")
        p_name = c1.text_input("Paid To *")
        p_amt = c2.number_input("Amount (₹) *", min_value=0.0, step=50.0)
        p_mode = c1.selectbox("Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"])
        p_desc = c2.text_input("Expense Purpose *")
        if st.button("💾 Save Expense to Cloud", type="primary", use_container_width=True):
            if p_name and p_amt > 0:
                GSheetsManager.append_row("Expense", {
                    "Date": v_date, "Expense Name": f"{p_name} ({p_desc})",
                    "Amount": p_amt, "Payment Mode": p_mode, "Notes": f"VOU #{v_no} | {p_mode}"
                })
                st.success("Expense Recorded in Google Sheets!")

    elif bill_type == "✏️ Edit / Delete Vouchers (Requires PIN)":
        st.markdown("### 🔐 Modify or Delete Payment Voucher Record")
        df_exp = GSheetsManager.get_df("Expense")
        
        if not df_exp.empty and "ID" in df_exp.columns:
            exp_clean = df_exp[df_exp["ID"].notna()].copy()
            exp_ids = exp_clean["ID"].tolist()
            exp_labels = [f"ID #{r['ID']} - {str(r.get('Expense Name', ''))} ({format_to_ddmmyyyy(r.get('Date', ''))}) | ₹{r.get('Amount', 0)} [{r.get('Payment Mode', 'Cash')}]" for _, r in exp_clean.iterrows()]
            
            chosen_e_idx = st.selectbox("Select Voucher / Expense to Modify:", range(len(exp_labels)), format_func=lambda i: exp_labels[i], key="edit_exp_select_idx")
            sel_exp_id = exp_ids[chosen_e_idx]
            exp_row = exp_clean[exp_clean["ID"].astype(str).str.replace(r'\.0$', '', regex=True) == str(sel_exp_id).replace('.0', '')].iloc[0]
            
            with st.expander(f"📝 Edit Expense Voucher #{sel_exp_id}", expanded=True):
                up_ed1, up_ed2 = st.columns(2)
                raw_exp_dt = parse_date_safely(exp_row.get("Date", ""))
                up_e_date = up_ed1.date_input("Date", raw_exp_dt, format="DD/MM/YYYY", key=f"e_dt_inp_{sel_exp_id}").strftime("%d/%m/%Y")
                up_e_name = up_ed2.text_input("Expense Description / Paid To", str(exp_row.get("Expense Name", "")), key=f"e_nm_{sel_exp_id}")
                
                up_ed3, up_ed4 = st.columns(2)
                up_e_amt = up_ed3.number_input("Amount (₹)", value=float(pd.to_numeric(exp_row.get("Amount", 0.0), errors='coerce') or 0.0), step=50.0, key=f"e_am_{sel_exp_id}")
                curr_exp_mode = str(exp_row.get("Payment Mode", "Cash"))
                exp_mode_idx = ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"].index(curr_exp_mode) if curr_exp_mode in ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"] else 0
                up_e_mode = up_ed4.selectbox("Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"], index=exp_mode_idx, key=f"e_mode_{sel_exp_id}")
                
                up_e_notes = st.text_input("Notes", str(exp_row.get("Notes", "")) if pd.notna(exp_row.get("Notes")) else "", key=f"e_nt_{sel_exp_id}")
                
                st.markdown("🔒 **Security Authorization:**")
                exp_auth_pin = st.text_input("Enter Security PIN to Authorize:", type="password", key=f"exp_pin_{sel_exp_id}")
                
                eb_up_col, eb_del_col = st.columns(2)
                if eb_up_col.button("🔄 Update Voucher", key=f"up_exp_btn_{sel_exp_id}", use_container_width=True):
                    if exp_auth_pin == GSheetsManager.get_pin():
                        GSheetsManager.update_row("Expense", sel_exp_id, {
                            "Date": str(up_e_date), "Expense Name": str(up_e_name),
                            "Amount": float(up_e_amt), "Payment Mode": str(up_e_mode), "Notes": str(up_e_notes)
                        })
                        st.success(f"✅ Voucher #{sel_exp_id} updated in Google Sheets!")
                        st.rerun()
                    else:
                        st.error("❌ Incorrect Security PIN!")
                        
                if eb_del_col.button("🗑️ Delete Voucher", key=f"del_exp_btn_{sel_exp_id}", type="primary", use_container_width=True):
                    if exp_auth_pin == GSheetsManager.get_pin():
                        GSheetsManager.delete_row("Expense", sel_exp_id)
                        st.warning(f"🗑️ Voucher #{sel_exp_id} deleted!")
                        st.rerun()
                    else:
                        st.error("❌ Incorrect Security PIN!")
        else:
            st.info("No expense vouchers found to edit.")

    elif bill_type == "Settle Old Pending Due":
        df_baki = GSheetsManager.get_df("Udhar_Baki")
        if not df_baki.empty and "Pending Amount" in df_baki.columns:
            pending = df_baki[pd.to_numeric(df_baki["Pending Amount"], errors='coerce') > 0]
            if not pending.empty:
                pending_ids = pending["ID"].tolist()
                pending_labels = [f"ID #{r['ID']} - {r['Customer Name']} | Pending: ₹{r['Pending Amount']}" for _, r in pending.iterrows()]
                
                sel_p_idx = st.selectbox("Select Due Account:", range(len(pending_labels)), format_func=lambda i: pending_labels[i])
                sel_id = pending_ids[sel_p_idx]
                
                r = pending[pending["ID"].astype(str).str.replace(r'\.0$', '', regex=True) == str(sel_id).replace('.0', '')].iloc[0]
                c_name = str(r.get('Customer Name', ''))
                c_serv = str(r.get('Service Details', 'Service'))
                curr_pend = float(pd.to_numeric(r.get('Pending Amount', 0), errors='coerce'))
                curr_paid = float(pd.to_numeric(r.get('Paid Amount', 0), errors='coerce'))
                
                s_amt = st.number_input("Payment Received Now (₹) *", min_value=0.0, max_value=curr_pend, value=curr_pend, step=100.0)
                s_mode = st.selectbox("Payment Mode", ["Cash", "UPI / GPay", "Bank Transfer", "Cheque"])
                if st.button("💳 Settle Balance", type="primary", use_container_width=True):
                    new_paid = curr_paid + s_amt
                    new_pending = curr_pend - s_amt
                    new_stat = "Cleared" if new_pending <= 0 else "Pending"
                    
                    GSheetsManager.update_row("Udhar_Baki", sel_id, {
                        "Paid Amount": float(new_paid),
                        "Pending Amount": float(new_pending),
                        "Status": str(new_stat)
                    })
                    GSheetsManager.append_row("Income", {
                        "Date": datetime.now().strftime("%d/%m/%Y"),
                        "Customer Person": str(c_name),
                        "Work Details": f"Due Settlement ({c_serv})",
                        "Amount": float(s_amt),
                        "Payment Mode": str(s_mode),
                        "Notes": f"Due Rec #{sel_id}"
                    })
                    st.success("Due Settled in Google Sheets!")
                    st.rerun()

# ----------------- 4. DUE COLLECTIONS -----------------
elif menu == "📋 Due Collections":
    st.subheader("📋 Due Collections & Credit Ledger Management (Cloud)")
    
    tab_due_view, tab_due_edit = st.tabs(["📋 Active Due Receivables", "🔐 Edit / Modify Due Record (Requires PIN)"])
    
    with tab_due_view:
        df_b = GSheetsManager.get_df("Udhar_Baki")
        if not df_b.empty:
            st.dataframe(df_b, use_container_width=True)
        else:
            st.info("No due records found.")
            
    with tab_due_edit:
        st.markdown("##### 🔐 Modify or Delete Customer Due Record")
        df_b = GSheetsManager.get_df("Udhar_Baki")
        if not df_b.empty and "ID" in df_b.columns:
            due_records_clean = df_b[df_b["ID"].notna()].copy()
            if not due_records_clean.empty:
                due_id_list = due_records_clean["ID"].tolist()
                due_display_list = [
                    f"ID #{r['ID']} - {str(r.get('Customer Name', ''))} ({clean_phone_number(r.get('Mobile Number', ''))}) | Pending: ₹{r.get('Pending Amount', 0)}" 
                    for _, r in due_records_clean.iterrows()
                ]
                
                chosen_idx = st.selectbox("Select Due Record to Edit / Delete:", range(len(due_display_list)), format_func=lambda i: due_display_list[i], key="due_mod_select_idx")
                sel_due_id = due_id_list[chosen_idx]
                due_r = due_records_clean[due_records_clean["ID"].astype(str).str.replace(r'\.0$', '', regex=True) == str(sel_due_id).replace('.0', '')].iloc[0]
                
                with st.expander(f"📝 Edit Due Record #{sel_due_id} - {due_r.get('Customer Name', '')}", expanded=True):
                    dc1, dc2 = st.columns(2)
                    raw_d_dt = parse_date_safely(due_r.get("Date", ""))
                    up_d_date = dc1.date_input("Entry Date", raw_d_dt, format="DD/MM/YYYY", key=f"d_dt_inp_{sel_due_id}").strftime("%d/%m/%Y")
                    up_d_name = dc2.text_input("Customer Name", str(due_r.get("Customer Name", "")), key=f"d_nm_{sel_due_id}")
                    
                    dc3, dc4 = st.columns(2)
                    up_d_phone = dc3.text_input("Mobile Number", clean_phone_number(due_r.get("Mobile Number", "")), key=f"d_ph_{sel_due_id}")
                    up_d_serv = dc4.text_input("Service Details", str(due_r.get("Service Details", "Service")), key=f"d_sv_{sel_due_id}")
                    
                    dc5, dc6 = st.columns(2)
                    up_d_tot = dc5.number_input("Total Amount (₹)", value=float(pd.to_numeric(due_r.get("Total Amount", 0.0), errors='coerce') or 0.0), step=100.0, key=f"d_tot_{sel_due_id}")
                    up_d_paid = dc6.number_input("Paid Amount (₹)", value=float(pd.to_numeric(due_r.get("Paid Amount", 0.0), errors='coerce') or 0.0), max_value=float(up_d_tot), step=100.0, key=f"d_pd_{sel_due_id}")
                    
                    up_d_pending = up_d_tot - up_d_paid
                    st.write(f"**Recalculated Pending Amount:** ₹ {up_d_pending:,.2f}")
                    
                    dc7, dc8 = st.columns(2)
                    raw_due_dt = parse_date_safely(due_r.get("Due Date", ""))
                    up_d_due_dt = dc7.date_input("Due Date", raw_due_dt, format="DD/MM/YYYY", key=f"d_ddt_inp_{sel_due_id}").strftime("%d/%m/%Y")
                    up_d_stat = dc8.selectbox("Status", ["Pending", "Cleared"], index=0 if up_d_pending > 0 else 1, key=f"d_st_{sel_due_id}")
                    
                    st.markdown("🔒 **Security Authorization:**")
                    due_auth_pin = st.text_input("Enter Security PIN to Authorize:", type="password", key=f"due_pin_{sel_due_id}")
                    
                    db_up_c, db_del_c = st.columns(2)
                    if db_up_c.button("🔄 Update Due Record", key=f"btn_up_due_{sel_due_id}", use_container_width=True):
                        if due_auth_pin == GSheetsManager.get_pin():
                            GSheetsManager.update_row("Udhar_Baki", sel_due_id, {
                                "Date": str(up_d_date), "Customer Name": str(up_d_name),
                                "Mobile Number": clean_phone_number(up_d_phone), "Service Details": str(up_d_serv),
                                "Total Amount": float(up_d_tot), "Paid Amount": float(up_d_paid),
                                "Pending Amount": float(up_d_pending), "Due Date": str(up_d_due_dt), "Status": str(up_d_stat)
                            })
                            st.success(f"✅ Due Record #{sel_due_id} updated in Google Sheets!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
                            
                    if db_del_c.button("🗑️ Delete Due Record", key=f"btn_del_due_{sel_due_id}", type="primary", use_container_width=True):
                        if due_auth_pin == GSheetsManager.get_pin():
                            GSheetsManager.delete_row("Udhar_Baki", sel_due_id)
                            st.warning(f"🗑️ Due Record #{sel_due_id} deleted from Google Sheets!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
        else:
            st.info("No due records found to edit.")

# ----------------- 5. REPORTS & PDF -----------------
elif menu == "📄 Reports & PDF":
    st.subheader("📄 Financial Reports & Statements")
    c1, c2 = st.columns(2)
    d_from = c1.date_input("From Date", datetime.now().replace(day=1), format="DD/MM/YYYY")
    d_to = c2.date_input("To Date", datetime.now(), format="DD/MM/YYYY")
    
    d_from_str = d_from.strftime("%d/%m/%Y")
    d_to_str = d_to.strftime("%d/%m/%Y")
    
    df_i = GSheetsManager.get_df("Income")
    df_e = GSheetsManager.get_df("Expense")
    df_b = GSheetsManager.get_df("Udhar_Baki")
    
    def filter_by_date_range(df, col_name, start_date, end_date):
        if df.empty or col_name not in df.columns:
            return pd.DataFrame()
        temp_df = df.copy()
        temp_df['_parsed_date'] = temp_df[col_name].apply(parse_date_safely)
        filtered = temp_df[(temp_df['_parsed_date'] >= start_date) & (temp_df['_parsed_date'] <= end_date)].drop(columns=['_parsed_date'])
        return filtered

    f_i = filter_by_date_range(df_i, "Date", d_from, d_to)
    f_e = filter_by_date_range(df_e, "Date", d_from, d_to)
    f_b = filter_by_date_range(df_b, "Date", d_from, d_to) if not df_b.empty else pd.DataFrame()
    
    t_i = pd.to_numeric(f_i["Amount"], errors='coerce').sum() if not f_i.empty and "Amount" in f_i.columns else 0.0
    t_e = pd.to_numeric(f_e["Amount"], errors='coerce').sum() if not f_e.empty and "Amount" in f_e.columns else 0.0
    cash_op, bank_op = GSheetsManager.get_opening_balance()
    tot_op = cash_op + bank_op
    closing_bal = tot_op + t_i - t_e
    
    st.info(f"**Period:** {d_from_str} to {d_to_str} | **Revenue:** ₹{t_i:,.2f} | **Expenses:** ₹{t_e:,.2f} | **Closing Balance:** ₹{closing_bal:,.2f}")
    
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
                    Paragraph(format_to_ddmmyyyy(r.get("Date", "-")), tbl_text),
                    Paragraph(str(r.get("Customer Person", "-")), tbl_text),
                    Paragraph(str(r.get("Work Details", "-")), tbl_text),
                    Paragraph(str(r.get("Payment Mode", "-")), tbl_text),
                    Paragraph(f"<b>{float(pd.to_numeric(r.get('Amount', 0), errors='coerce')):,.2f}</b>", tbl_text)
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
                Paragraph("Mode", tbl_hdr),
                Paragraph("Notes / Remarks", tbl_hdr), 
                Paragraph("Amount (Rs.)", tbl_hdr)
            ]]
            for _, r in df_exp.iterrows():
                e_rows.append([
                    Paragraph(format_to_ddmmyyyy(r.get("Date", "-")), tbl_text),
                    Paragraph(str(r.get("Expense Name", "-")), tbl_text),
                    Paragraph(str(r.get("Payment Mode", "Cash")), tbl_text),
                    Paragraph(str(r.get("Notes", "-")), tbl_text),
                    Paragraph(f"<b>{float(pd.to_numeric(r.get('Amount', 0), errors='coerce')):,.2f}</b>", tbl_text)
                ])
            t2 = Table(e_rows, colWidths=[80, 300, 80, 230, 100])
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
            active_dues = df_due[pd.to_numeric(df_due["Pending Amount"], errors='coerce') > 0] if "Pending Amount" in df_due.columns else pd.DataFrame()
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
                        Paragraph(format_to_ddmmyyyy(r.get("Date", "-")), tbl_text),
                        Paragraph(str(r.get("Customer Name", "-")), tbl_text),
                        Paragraph(clean_phone_number(r.get("Mobile Number", "-")), tbl_text),
                        Paragraph(str(r.get("Service Details", "Service")), tbl_text),
                        Paragraph(f"{float(pd.to_numeric(r.get('Total Amount', 0), errors='coerce')):,.2f}", tbl_text),
                        Paragraph(f"{float(pd.to_numeric(r.get('Paid Amount', 0), errors='coerce')):,.2f}", tbl_text),
                        Paragraph(f"<b>{float(pd.to_numeric(r.get('Pending Amount', 0), errors='coerce')):,.2f}</b>", tbl_text),
                        Paragraph(format_to_ddmmyyyy(r.get("Due Date", "-")), tbl_text)
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

    stat_pdf = generate_statement_pdf_landscape(d_from_str, d_to_str, f_i, f_e, f_b, tot_op, t_i, t_e, closing_bal)
    st.download_button(
        label="📥 Download Complete Financial Statement (Landscape PDF)",
        data=stat_pdf,
        file_name=f"Landscape_Financial_Statement_{d_from_str.replace('/', '-')}_to_{d_to_str.replace('/', '-')}.pdf",
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
    st.subheader("🏦 Opening Balance Setup (Google Sheets)")
    curr_c, curr_b = GSheetsManager.get_opening_balance()
    c1, c2 = st.columns(2)
    in_c = c1.number_input("Cash in Hand (₹)", value=float(curr_c), step=500.0)
    in_b = c2.number_input("Bank Balance (₹)", value=float(curr_b), step=500.0)
    pin = st.text_input("Enter Security PIN to Save:", type="password")
    if st.button("💾 Save Opening Balance", type="primary", use_container_width=True):
        if pin == GSheetsManager.get_pin():
            GSheetsManager.set_opening_balance(in_c, in_b)
            st.success("Opening Balance Saved to Google Sheets!")
            st.rerun()
        else:
            st.error("Invalid PIN!")

# ----------------- 7. CUSTOMERS DIRECTORY -----------------
elif menu == "👥 Customers Directory":
    st.subheader("👥 Client Directory & Broadcast (Cloud)")
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
                    df_c = GSheetsManager.get_df("Customers")
                    clean_phone = clean_phone_number(cp)
                    if not df_c.empty and "Mobile Number" in df_c.columns and clean_phone in df_c["Mobile Number"].astype(str).values:
                        st.warning(f"⚠️ A customer with mobile {clean_phone} already exists in records!")
                    else:
                        GSheetsManager.append_row("Customers", {
                            "Created Date": datetime.now().strftime("%d/%m/%Y"),
                            "Customer Name": cn.strip(),
                            "Mobile Number": clean_phone,
                            "City Address": c_addr,
                            "Primary Service": cs,
                            "Notes": c_notes
                        })
                        st.success(f"Client '{cn}' saved to Google Sheets successfully!")
                        st.rerun()
                else:
                    st.error("Customer name and mobile number are required.")
                    
    with tab_list:
        df_c = GSheetsManager.get_df("Customers")
        if not df_c.empty and "ID" in df_c.columns:
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
            
            clean_cust_rows = df_c[df_c["ID"].notna()].copy()
            if not clean_cust_rows.empty:
                cust_ids = clean_cust_rows["ID"].tolist()
                cust_labels = [
                    f"ID #{r['ID']} - {str(r.get('Customer Name', ''))} ({clean_phone_number(r.get('Mobile Number', ''))})" 
                    for _, r in clean_cust_rows.iterrows()
                ]
                
                selected_idx = st.selectbox(
                    "Select Customer to Edit / Update / Delete:", 
                    range(len(cust_labels)), 
                    format_func=lambda idx: cust_labels[idx],
                    key="cust_edit_select_box"
                )
                
                sel_c_id = cust_ids[selected_idx]
                c_row = clean_cust_rows[clean_cust_rows["ID"].astype(str).str.replace(r'\.0$', '', regex=True) == str(sel_c_id).replace('.0', '')].iloc[0]
                
                with st.expander(f"📝 Edit Client Profile #{sel_c_id} - {c_row.get('Customer Name', '')}", expanded=True):
                    ec1, ec2 = st.columns(2)
                    u_cname = ec1.text_input("Customer Name *", str(c_row.get("Customer Name", "")))
                    u_cphone = ec2.text_input("Mobile Number *", clean_phone_number(c_row.get("Mobile Number", "")))
                    
                    ec3, ec4 = st.columns(2)
                    u_caddr = ec3.text_input("Address / City / Village", str(c_row.get("City Address", "Kadi")) if pd.notna(c_row.get("City Address")) else "Kadi")
                    
                    curr_serv = str(c_row.get("Primary Service", "VISA"))
                    serv_idx = SERVICE_OPTIONS.index(curr_serv) if curr_serv in SERVICE_OPTIONS else 0
                    u_cserv = ec4.selectbox("Primary Service", SERVICE_OPTIONS, index=serv_idx)
                    
                    u_cnotes = st.text_area("Notes / Remarks", str(c_row.get("Notes", "")) if pd.notna(c_row.get("Notes")) else "")
                    
                    st.markdown("🔒 **Security Confirmation:**")
                    edit_pin = st.text_input("Enter Security PIN:", type="password", key=f"c_pin_{sel_c_id}")
                    
                    b_col1, b_col2 = st.columns(2)
                    if b_col1.button("🔄 Update Customer Details", key=f"btn_up_{sel_c_id}", use_container_width=True):
                        if edit_pin == GSheetsManager.get_pin():
                            GSheetsManager.update_row("Customers", sel_c_id, {
                                "Customer Name": str(u_cname), "Mobile Number": clean_phone_number(u_cphone),
                                "City Address": str(u_caddr), "Primary Service": str(u_cserv), "Notes": str(u_cnotes)
                            })
                            st.success("Client profile updated successfully in Google Sheets!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
                            
                    if b_col2.button("🗑️ Delete Customer", key=f"btn_del_{sel_c_id}", type="primary", use_container_width=True):
                        if edit_pin == GSheetsManager.get_pin():
                            GSheetsManager.delete_row("Customers", sel_c_id)
                            st.warning("Client deleted from Google Sheets!")
                            st.rerun()
                        else:
                            st.error("❌ Incorrect Security PIN!")
        else:
            st.info("No registered clients found.")

    with tab_promo:
        st.markdown("##### 📢 Bulk Broadcast & Marketing Campaign")
        df_c = GSheetsManager.get_df("Customers")
        if not df_c.empty:
            serv_col_name = "Primary Service" if "Primary Service" in df_c.columns else df_c.columns[min(5, len(df_c.columns)-1)]
            service_unique_list = [str(x) for x in df_c[serv_col_name].dropna().unique() if str(x).strip() != ""] if serv_col_name in df_c.columns else []
            
            sel_aud = st.selectbox("Select Target Audience:", ["All Clients"] + service_unique_list, key="broadcast_aud_sel")
            target_df = df_c if (sel_aud == "All Clients" or serv_col_name not in df_c.columns) else df_c[df_c[serv_col_name].astype(str) == str(sel_aud)]
            
            st.write(f"**Total Recipients:** {len(target_df)}")
            display_cols = [c for c in ["Customer Name", "Mobile Number", "City Address", "Primary Service", "Notes"] if c in target_df.columns]
            st.dataframe(target_df[display_cols], use_container_width=True)
            
            promo_msg = st.text_area("Broadcast Message Template:", value=f"Greetings from {COMPANY_NAME}! Contact us at {COMPANY_MOBILE} for special offers and updates regarding your service inquiry.\n\nPayment Mobile No: {PAYMENT_MOBILE} | UPI ID: {UPI_ID}")
            
            st.markdown("##### 📲 Click to Send WhatsApp Directly:")
            for _, prow in target_df.head(25).iterrows():
                p_phone = clean_phone_number(prow.get("Mobile Number", ""))
                p_name = prow.get("Customer Name", "Client")
                p_city = prow.get("City Address", "Kadi")
                if p_phone and len(p_phone) >= 10:
                    p_url = f"https://wa.me/91{p_phone}?text={urllib.parse.quote(promo_msg)}"
                    st.markdown(f"👉 **{p_name}** ({p_phone}) - [{p_city}]: [📲 Send WhatsApp]({p_url})")
        else:
            st.info("No client records available for marketing broadcast.")

# ----------------- 8. INCOME & EXPENSE MANAGEMENT -----------------
elif menu == "💰 Income":
    st.subheader("💰 Income Ledger (Google Sheets)")
    df_i = GSheetsManager.get_df("Income")
    if not df_i.empty:
        st.dataframe(df_i, use_container_width=True)

elif menu == "💸 Expenses":
    st.subheader("💸 Expense Ledger (Google Sheets)")
    df_e = GSheetsManager.get_df("Expense")
    if not df_e.empty:
        st.dataframe(df_e, use_container_width=True)

# ----------------- 9. CLOUD BACKUP -----------------
elif menu == "💾 Cloud Excel Backup":
    st.subheader("💾 Complete Data Export (Excel Workbook)")
    st.info("💡 Generate an offline Excel backup file (.xlsx) containing all Google Sheets data.")
    
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for tbl in ["Customers", "Invoices_Archive", "Income", "Expense", "Udhar_Baki", "Task_Reminder", "Settings"]:
            df = GSheetsManager.get_df(tbl)
            df.to_excel(writer, sheet_name=tbl, index=False)
    buf.seek(0)
    
    st.download_button(
        label="📥 Download Full Offline Excel Backup (.xlsx)",
        data=buf,
        file_name=f"Rojmed_Cloud_Backup_{datetime.now().strftime('%d_%m_%Y_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

# ----------------- 10. SECURITY PIN -----------------
elif menu == "⚙️ Security / Change PIN":
    st.subheader("⚙️ Change Master PIN")
    with st.form("pin_form"):
        old_p = st.text_input("Current PIN *", type="password")
        new_p = st.text_input("New PIN *", type="password")
        conf_p = st.text_input("Confirm New PIN *", type="password")
        if st.form_submit_button("💾 Update PIN"):
            if old_p == GSheetsManager.get_pin():
                if new_p and new_p == conf_p:
                    GSheetsManager.set_pin(new_p)
                    st.success("PIN Updated in Google Sheets Settings!")
                else:
                    st.error("PIN mismatch!")
            else:
                st.error("Incorrect Current PIN!")
