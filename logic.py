# logic.py

import os
import time
import datetime
from datetime import date, timedelta
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage
#import win32com.client
#import pythoncom
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit as st
import json
from io import BytesIO
import fitz # PyMuPDF for editing the credit card form
import urllib.parse
import resend
from fpdf import FPDF
#import fitz

# =============================================================================
# CONSTANTS
# =============================================================================
GSHEET_NAME = "Estimate form"
ESTIMATE_SHEET_NAME = "Estimate Form MOAS"
PRICE_LIBRARY_SHEET_NAME = "Price Library"
HISTORY_SHEET_NAME = "S/N EMAIL history"
MAIN_DATA_SHEET_INDEX = 1
CREDS_FILE = "Credentials.json"
ARCHIVE_SHEET_NAME = "DailyReportArchive"
EOD_SUMMARY_ARCHIVE_SHEET_NAME = "EODSummaryArchive"
ARCHIVE_SHEET_HEADERS = ["Report Date", "Needs Estimate Creation", "Needs Shipping", "Needs Reminder"]
EOD_ARCHIVE_SHEET_HEADERS = ["Report Date", "Estimate Task Summary", "Reminder Task Summary", "Shipping Task Summary", "AdHoc Shipped Today"]
EXPECTED_COLUMN_ORDER = [
    "RMA", "SPC Code", "Part Number", "S/N", "Description", "Fault Comments", "Resolution Comments", "Sender",
    "Estimate Complete Time", "Estimate Complete", "Estimate Approved", "Estimate Approved Time",
    "Estimate Sent To Email", "Estimate Sent Time", "Reminder Completed", "Reminder Completed Time", "Reminder Contact Method",
    "QA Approved", "QA Approved Time", "Shipped", "Shipped Time"
]
BC_BASE_URL = "https://businesscentral.dynamics.com/7bcfb5b0-27a1-4e18-99d8-ca66570addd8/Production"
BC_COMPANY = "PROD"
BC_PAGE_ID = "70001"
BC_RMA_FIELD_NAME = "No."
BC_LINK_COL_NAME = "View in BC"
SOURCE_PARTS_ARCHIVE_DIR = "source_parts_archive" # <-- ADDED: New constant for the archive folder

# =============================================================================
# CORE GOOGLE SHEETS & DATA LOADING
# =============================================================================
@st.cache_resource
def connect_to_google_sheet():
    '''Connects to Google Sheets using service account credentials.'''
    try:
        # 'scopes' is defined here
        scopes = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                  "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scopes) # and used here
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Failed to connect to Google Sheets: {e}")
        return None

@st.cache_data(ttl=300)
def load_data_from_google_sheet():
    '''Loads and preprocesses data from the main Google Sheet.'''
    client = connect_to_google_sheet()
    if client is None: return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER)
    try:
        worksheet = client.open(GSHEET_NAME).get_worksheet(MAIN_DATA_SHEET_INDEX)
        all_values = worksheet.get_all_values()
        if not all_values: return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER)

        headers_from_sheet = all_values[0]
        temp_df = pd.DataFrame(all_values[1:], columns=headers_from_sheet)
        df = pd.DataFrame(columns=EXPECTED_COLUMN_ORDER)

        for col in EXPECTED_COLUMN_ORDER:
            if col in temp_df.columns:
                df[col] = temp_df[col]
            else:
                df[col] = pd.NaT if "Time" in col else "N/A"

        for col in df.columns:
            if "Time" in col:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            else:
                df[col] = df[col].astype(str).replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'N/A')
                if col in ["Estimate Complete", "Estimate Approved", "Reminder Completed", "QA Approved", "Shipped"]:
                    df[col] = df[col].replace('N/A', 'No')

        return df[EXPECTED_COLUMN_ORDER]
    except Exception as e:
        st.error(f"An error occurred loading data: {e}")
        return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER)

def gsheet_update_wrapper(update_function, target_sheet_name, *args):
    '''A wrapper to handle GSheet connection for update operations.'''
    client = connect_to_google_sheet()
    if not client: return False
    try:
        worksheet = client.open(GSHEET_NAME).worksheet(target_sheet_name)
        headers = worksheet.row_values(1)
        return update_function(worksheet, headers, *args)
    except gspread.exceptions.WorksheetNotFound:
        st.warning(f"Worksheet '{target_sheet_name}' not found. It may be created if needed.")
        try:
             worksheet = client.open(GSHEET_NAME).worksheet(target_sheet_name)
             headers = []
             return update_function(worksheet, headers, *args)
        except Exception as e_inner:
             st.error(f"GSheet Update Error on sheet '{target_sheet_name}': {e_inner}")
             return False

    except Exception as e:
        st.error(f"GSheet Update Error on sheet '{target_sheet_name}': {e}")
        return False

def find_row_in_gsheet(worksheet, search_rma, search_sn, headers):
    '''Finds a specific row by RMA and S/N.'''
    try:
        rma_col_idx = headers.index("RMA")
        sn_col_idx = headers.index("S/N")
    except ValueError: return -1
    search_by_sn_only = str(search_rma).strip().lower() in ['n/a', '']
    all_values = worksheet.get_all_values()
    for i, row in enumerate(all_values[1:], start=2):
        rma_val = str(row[rma_col_idx]).strip().lower() if len(row) > rma_col_idx else ""
        sn_val = str(row[sn_col_idx]).strip().lower() if len(row) > sn_col_idx else ""
        if search_by_sn_only:
            if sn_val == str(search_sn).strip().lower() and rma_val in ['n/a', '']:
                return i
        elif rma_val == str(search_rma).strip().lower() and sn_val == str(search_sn).strip().lower():
            return i
    return -1

# =============================================================================
# PRICE LIBRARY LOGIC
# =============================================================================
@st.cache_data(ttl=600)
def load_price_library():
    '''Loads the price library from Google Sheets into a dictionary.'''
    client = connect_to_google_sheet()
    if not client:
        return {}
    try:
        spreadsheet = client.open(GSHEET_NAME)
        worksheet = spreadsheet.worksheet(PRICE_LIBRARY_SHEET_NAME)
        records = worksheet.get_all_records()
        price_map = {str(rec.get('No.')): rec.get('Amount Including Tax') for rec in records if 'No.' in rec and 'Amount Including Tax' in rec}
        return price_map
    except gspread.exceptions.WorksheetNotFound:
        st.info("Price Library sheet not found. It will be created when a new estimate is generated.")
        return {}
    except Exception as e:
        st.error(f"Error loading Price Library: {e}")
        return {}

def update_price_library_and_usage_count(parts_df):
    '''
    Adds/updates part prices, increments usage count, and saves part descriptions.
    '''
    client = connect_to_google_sheet()
    if not client:
        return False
    try:
        spreadsheet = client.open(GSHEET_NAME)
        try:
            worksheet = spreadsheet.worksheet(PRICE_LIBRARY_SHEET_NAME)
            headers = worksheet.row_values(1)
            if "Usage Count" not in headers:
                worksheet.update_cell(1, len(headers) + 1, "Usage Count")
                headers.append("Usage Count")

            library_data = worksheet.get_all_records()
            part_map = {
                str(rec.get('No.')): {
                    'row': i + 2,
                    'count': int(rec.get('Usage Count') or 0),
                } for i, rec in enumerate(library_data)
            }
            existing_part_numbers = set(part_map.keys())

        except gspread.exceptions.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title=PRICE_LIBRARY_SHEET_NAME, rows="1000", cols=4)
            headers = ["No.", "Description", "Amount Including Tax", "Usage Count"]
            worksheet.append_row(headers)
            part_map = {}
            existing_part_numbers = set()

        updates = []
        parts_in_estimate = parts_df['No.'].astype(str).tolist()
        usage_col_index = headers.index("Usage Count") + 1

        for part_no in parts_in_estimate:
            if part_no in part_map:
                row_index = part_map[part_no]['row']
                new_count = part_map[part_no]['count'] + 1
                part_map[part_no]['count'] = new_count
                updates.append({
                    'range': gspread.utils.rowcol_to_a1(row_index, usage_col_index),
                    'values': [[new_count]],
                })

        if updates:
            worksheet.batch_update(updates)

        excluded_parts = ['BILLABLE FREIGHT', 'TECHNICIAN HQ']
        new_parts_df = parts_df[
            (~parts_df['No.'].astype(str).isin(existing_part_numbers)) &
            (parts_df['Amount Including Tax'].notna()) &
            (parts_df['Amount Including Tax'] > 0) &
            (~parts_df['No.'].isin(excluded_parts))
        ].copy()

        if not new_parts_df.empty:
            new_parts_df.loc[:, 'Usage Count'] = 1
            new_rows = new_parts_df[['No.', 'Description', 'Amount Including Tax', 'Usage Count']].values.tolist()
            worksheet.append_rows(new_rows, value_input_option='USER_ENTERED')
            st.info(f"Added {len(new_rows)} new part(s) to the Price Library.")

        return True
    except Exception as e:
        st.error(f"Failed to update Price Library & Usage Count: {e}")
        return False

# =============================================================================
# ACTION LOGIC (UPDATE GOOGLE SHEET)
# =============================================================================
def _add_or_update_estimate_in_sheet(worksheet, headers, form_data, parts_df):
    rma = form_data.get('rma')
    sn = form_data.get('serial')
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total_sum = float(parts_df['Line Total'].sum()) if 'Line Total' in parts_df.columns else 0.0

    parts_json = parts_df.to_json(orient='records')

    if "Parts JSON" not in headers:
        worksheet.update_cell(1, len(headers) + 1, "Parts JSON")

    row_data = [
        rma, sn, form_data.get('contact'), form_data.get('cust_name'),
        form_data.get('cust_num'), timestamp, total_sum,
        form_data.get('description'), form_data.get('evaluation'), parts_json
    ]
    worksheet.append_row(row_data)
    return True

def add_or_update_estimate_in_gsheet(form_data, parts_df):
    return gsheet_update_wrapper(_add_or_update_estimate_in_sheet, ESTIMATE_SHEET_NAME, form_data, parts_df)

def _update_estimate_sent_in_sheet(worksheet, headers, rma, sn, email, sent_date):
    row = find_row_in_gsheet(worksheet, rma, sn, headers)
    if row != -1:
        ts = datetime.datetime.combine(sent_date, datetime.datetime.now().time()).strftime("%Y-%m-%d %H:%M:%S")
        updates = [{'range': f'M{row}', 'values': [[email]]}, {'range': f'N{row}', 'values': [[ts]]}]
        worksheet.batch_update(updates)
        return True
    return False

def update_estimate_sent_details_in_gsheet(rma, sn, email, sent_date):
    client = connect_to_google_sheet()
    if not client: return False
    try:
        worksheet = client.open(GSHEET_NAME).get_worksheet(MAIN_DATA_SHEET_INDEX)
        headers = worksheet.row_values(1)
        return _update_estimate_sent_in_sheet(worksheet, headers, rma, sn, email, sent_date)
    except Exception as e:
        st.error(f"Error updating estimate sent details: {e}")
        return False

def _update_shipped_in_sheet(worksheet, headers, rma, sn, shipped_date):
    row = find_row_in_gsheet(worksheet, rma, sn, headers)
    if row != -1:
        ts = datetime.datetime.combine(shipped_date, datetime.datetime.now().time()).strftime("%Y-%m-%d %H:%M:%S")
        worksheet.batch_update([{'range': f'T{row}', 'values': [['Yes']]}, {'range': f'U{row}', 'values': [[ts]]}])
        return True
    return False

def update_shipped_status_in_gsheet(rma, sn, shipped_date):
    return gsheet_update_wrapper(_update_shipped_in_sheet, HISTORY_SHEET_NAME, rma, sn, shipped_date)

def _update_reminder_in_sheet(worksheet, headers, rma, sn, reminder_date, method):
    row = find_row_in_gsheet(worksheet, rma, sn, headers)
    if row != -1:
        ts = datetime.datetime.combine(reminder_date, datetime.datetime.now().time()).strftime("%Y-%m-%d %H:%M:%S")
        worksheet.batch_update([{'range': f'O{row}', 'values': [['Yes']]}, {'range': f'P{row}', 'values': [[ts]]}, {'range': f'Q{row}', 'values': [[method]]}])
        return True
    return False

def update_reminder_details_in_gsheet(rma, sn, reminder_date, method):
    return gsheet_update_wrapper(_update_reminder_in_sheet, HISTORY_SHEET_NAME, rma, sn, reminder_date, method)



# =============================================================================
# FILE GENERATION & EMAIL
# =============================================================================
def generate_estimate_files(form_data, parts_df, save_directory):
    try:
        template_path = "estimate_form_template.pdf"
        if not os.path.exists(template_path):
            st.error(f"Template file not found at {template_path}. Please create and upload the fillable PDF template.")
            return None

        # --- File Saving Path ---
        sanitized_rma = "".join(c for c in form_data.get('rma', 'file') if c.isalnum() or c in ('_')).rstrip()
        file_name_base = os.path.join(save_directory, f"Estimate_Form_{sanitized_rma}")
        pdf_path = f"{file_name_base}.pdf"

        # --- Open the PDF Template ---
        doc = fitz.open(template_path)
        
        # --- Fill in the Main Fields ---
        field_map = {
            "rma": str(form_data.get('rma', '')),
            "serial": str(form_data.get('serial', '')),
            "contact": str(form_data.get('contact', '')),
            "cust_name": str(form_data.get('cust_name', '')),
            "description": str(form_data.get('description', '')),
            "evaluation": str(form_data.get('evaluation', ''))
        }
        
        for page in doc:
            for field_name, field_value in field_map.items():
                field = page.load_widget_by_name(field_name)
                if field:
                    field.field_value = field_value
                    field.update()

        # --- Fill in the Parts Table (assumes up to 13 rows) ---
        total_cost = 0
        for index, row in parts_df.head(13).iterrows(): # .head(13) to prevent errors if more parts
            line_total = float(row.get('Quantity', 1)) * float(row.get('Amount Including Tax', 0))
            total_cost += line_total
            
            # Find and fill each widget in the row
            for page in doc:
                # Part Number
                no_field = page.load_widget_by_name(f"part_no_{index}")
                if no_field: no_field.field_value = str(row['No.']); no_field.update()
                # Description
                desc_field = page.load_widget_by_name(f"part_desc_{index}")
                if desc_field: desc_field.field_value = str(row['Description']); desc_field.update()
                # Quantity
                qty_field = page.load_widget_by_name(f"part_qty_{index}")
                if qty_field: qty_field.field_value = str(row['Quantity']); qty_field.update()
                # Unit Price
                price_field = page.load_widget_by_name(f"part_price_{index}")
                if price_field: price_field.field_value = f"${float(row['Amount Including Tax']):.2f}"; price_field.update()
                # Line Total
                total_field = page.load_widget_by_name(f"part_total_{index}")
                if total_field: total_field.field_value = f"${line_total:.2f}"; total_field.update()
        
        # --- Fill in the Final Total ---
        for page in doc:
            final_total_field = page.load_widget_by_name("final_total")
            if final_total_field:
                final_total_field.field_value = f"${total_cost:.2f}"
                final_total_field.update()

        # Save the filled PDF
        doc.save(pdf_path, garbage=3, deflate=True, clean=True)
        doc.close()

        return {'excel_path': None, 'pdf_path': pdf_path}
    except Exception as e:
        st.error(f"An error occurred while filling the PDF template: {e}")
        return None

def send_estimate_email(recipient_email, rma_number, serial_number, estimate_pdf_path):
    """
    Sends the estimate PDF and credit card form using the Resend API.
    """
    try:
        # Initialize Resend with the API key from Streamlit secrets
        resend.api_key = st.secrets["resend"]["api_key"]

        # 1. Read the Estimate PDF file
        with open(estimate_pdf_path, "rb") as f:
            estimate_pdf_content = f.read()

        # 2. Find and read the Credit Card Form PDF
        cc_form_path = 'creditform/Credit_card_form.pdf'
        attachments_list = [
            {"filename": os.path.basename(estimate_pdf_path), "content": estimate_pdf_content}
        ]
        if os.path.exists(cc_form_path):
            with open(cc_form_path, "rb") as f:
                cc_form_content = f.read()
            attachments_list.append(
                {"filename": os.path.basename(cc_form_path), "content": cc_form_content}
            )

        # 3. Create the email body
        email_html_body = f"""
        <p>Greeting,</p>
        <p>Please review the estimate form that is attached to this email for S/N: <strong>{serial_number}</strong> and RMA: <strong>{rma_number}</strong>.</p>
        <p>If approved, please sign and send back the estimate to the following email: serviceorders@iridex.com. If paying by CC, fill out the attached credit card form and email it back. If paying by PO, please provide a hard copy of the PO.</p>
        <p>Finally, please confirm your shipping address to make sure we ship it to you with no issues. If you have any questions, please let us know.</p>
        """

        # 4. Send the email using the Resend API
        params = {
            "from": "Your App Name <onboarding@resend.dev>", # You can customize the "from" name
            "to": [recipient_email],
            "subject": f"Iridex's Estimate Form for S/N: {serial_number}, RMA: {rma_number}",
            "html": email_html_body,
            "attachments": attachments_list,
        }

        email = resend.Emails.send(params)
        st.success(f"Email sent successfully to {recipient_email}!")
        return True, cc_form_path

    except Exception as e:
        st.error(f"Failed to send email using Resend: {e}")
        return False, None
# =============================================================================
# SEARCH, REPORTING & ARCHIVING
# =============================================================================
def search_estimates_in_sheet(query, df):
    if df.empty or query == "": return pd.DataFrame()
    return df[df['RMA'].str.contains(query, case=False, na=False) | df['S/N'].str.contains(query, case=False, na=False)]

def identify_overdue_estimates(df, days_threshold=3):
    if df.empty: return pd.DataFrame()
    now = datetime.datetime.now()
    overdue_df = df[(df['Estimate Complete'].str.lower() == 'yes') & (df['Estimate Sent To Email'].str.lower() == 'n/a') & (df['Shipped'].str.lower().isin(['no', 'n/a'])) & (df['Estimate Complete Time'].notna()) & ((now - df['Estimate Complete Time']).dt.days > days_threshold)].copy()
    if not overdue_df.empty:
        overdue_df['Days Overdue for Sending'] = (now - overdue_df['Estimate Complete Time']).dt.days
        overdue_df[BC_LINK_COL_NAME] = overdue_df.apply(lambda r: f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(str(r['RMA']))}%27", axis=1)
    return overdue_df

def identify_overdue_reminders(df, days_threshold=2):
    if df.empty: return pd.DataFrame()
    now = datetime.datetime.now()
    overdue_df = df[(df['Estimate Sent To Email'].str.lower() != 'n/a') & (df['Reminder Completed'].str.lower().isin(['no', 'n/a'])) & (df['Estimate Approved'].str.lower().isin(['no', 'n/a'])) & (df['Estimate Sent Time'].notna()) & ((now - df['Estimate Sent Time']).dt.days > days_threshold)].copy()
    if not overdue_df.empty:
        overdue_df['Days Pending Reminder'] = (now - overdue_df['Estimate Sent Time']).dt.days
        overdue_df[BC_LINK_COL_NAME] = overdue_df.apply(lambda r: f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(str(r['RMA']))}%27", axis=1)
    return overdue_df

def identify_overdue_for_shipping(df, days_threshold=1):
    if df.empty: return pd.DataFrame()
    now = datetime.datetime.now()
    overdue_df = df[(df['Estimate Approved'].str.lower() == 'yes') & (df['QA Approved'].str.lower() == 'yes') & (df['Shipped'].str.lower().isin(['no', 'n/a'])) & (df['QA Approved Time'].notna()) & ((now - df['QA Approved Time']).dt.days > days_threshold)].copy()
    if not overdue_df.empty:
        overdue_df['Days Pending Shipping'] = (now - overdue_df['QA Approved Time']).dt.days
        overdue_df[BC_LINK_COL_NAME] = overdue_df.apply(lambda r: f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(str(r['RMA']))}%27", axis=1)
    return overdue_df

def generate_single_day_report_content(df, report_date_obj):
    report_content = { "date": report_date_obj.strftime("%Y-%m-%d"), "needs_shipping": [], "needs_estimate_creation": [], "needs_reminder": [] }
    shipping_df = df[(df['Estimate Complete'].astype(str).str.lower() == 'yes') & (df['Estimate Approved'].astype(str).str.lower() == 'yes') & (df['QA Approved'].astype(str).str.lower() == 'yes') & (df['Shipped'].astype(str).str.lower().isin(['no', 'n/a'])) & (pd.to_datetime(df['QA Approved Time'], errors='coerce').dt.date == report_date_obj) ]
    for _, row in shipping_df.iterrows():
        report_content["needs_shipping"].append({'RMA': str(row['RMA']), 'S/N': str(row['S/N']), 'SPC Code': str(row.get('SPC Code', 'N/A'))})
    day_prior_to_report = report_date_obj - timedelta(days=1)
    estimate_df = df[(df['Estimate Complete'].astype(str).str.lower() == 'yes') & (df['Estimate Sent To Email'].astype(str).str.lower() == 'n/a') & (pd.to_datetime(df['Estimate Complete Time'], errors='coerce').dt.date == day_prior_to_report) ]
    for _, row in estimate_df.iterrows():
        report_content["needs_estimate_creation"].append({'RMA': str(row['RMA']), 'S/N': str(row['S/N']), 'SPC Code': str(row.get('SPC Code', 'N/A')), 'Est. Complete Date': day_prior_to_report.strftime('%Y-%m-%d')})
    estimate_sent_target_date = report_date_obj - timedelta(days=2)
    reminder_df = df[(df['Estimate Sent To Email'].astype(str).str.lower() != 'n/a') & (df['Reminder Completed'].astype(str).str.lower().isin(['no', 'n/a'])) & (df['Estimate Approved'].astype(str).str.lower().isin(['no', 'n/a'])) & (pd.to_datetime(df['Estimate Sent Time'], errors='coerce').dt.date == estimate_sent_target_date)]
    for _, row in reminder_df.iterrows():
        report_content["needs_reminder"].append({'RMA': str(row['RMA']), 'S/N': str(row['S/N']),'SPC Code': str(row.get('SPC Code', 'N/A')), 'Estimate Sent To Email': str(row['Estimate Sent To Email']), 'Estimate Sent Time': pd.to_datetime(row['Estimate Sent Time']).strftime('%Y-%m-%d') if pd.notna(row['Estimate Sent Time']) else 'N/A'})
    return report_content

def get_archived_reports(archive_sheet_name):
    client = connect_to_google_sheet()
    if not client: return []
    try:
        worksheet = client.open(GSHEET_NAME).worksheet(archive_sheet_name)
        records = worksheet.get_all_records()
        for rec in records:
            for key, val in rec.items():
                if isinstance(val, str) and val.startswith('['):
                    try: rec[key] = json.loads(val)
                    except json.JSONDecodeError: rec[key] = []
        return sorted(records, key=lambda r: r.get('Report Date', ''), reverse=True)
    except gspread.exceptions.WorksheetNotFound: return []
    except Exception as e: st.error(f"Error loading archive '{archive_sheet_name}': {e}"); return []

def get_last_report_date_from_archive(archived_reports):
    if not archived_reports: return date.today() - timedelta(days=1)
    try:
        return datetime.datetime.strptime(archived_reports[0]['Report Date'], "%Y-%m-%d").date()
    except: return date.today() - timedelta(days=1)

def save_report_to_archive(report_data, archive_sheet_name, archive_headers):
    client = connect_to_google_sheet()
    if not client: return False
    try:
        spreadsheet = client.open(GSHEET_NAME)
        try:
            worksheet = spreadsheet.worksheet(archive_sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title=archive_sheet_name, rows="100", cols=len(archive_headers))
            worksheet.append_row(archive_headers)

        if report_data['date'] in worksheet.col_values(1): return False

        row_to_append = [report_data.get('date')]
        for header in archive_headers[1:]:
            row_to_append.append(json.dumps(report_data.get(header.replace(" ", "_").lower(), [])))
        worksheet.append_row(row_to_append)
        return True
    except Exception as e:
        st.error(f"Error saving report to archive '{archive_sheet_name}': {e}")
        return False

def create_excel_report_bytes(report_data, report_type="Daily"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        header_format = workbook.add_format({'bold': True, 'text_wrap': False, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
        cell_format = workbook.add_format({'border': 1})
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
        report_date_for_title = report_data.get('date', 'Unknown_Date')

        if report_type in ["Newly Generated", "Custom Date Report", "Archived Daily Report"]:
            sheets_data = { "Needs Estimate Creation": report_data.get("needs_estimate_creation", []), "Needs Reminder": report_data.get("needs_reminder", []), "Needs Shipping": report_data.get("needs_shipping", []) }
        elif report_type in ["EOD", "Archived EOD Summary"]:
            sheets_data = { "EOD Estimate Tasks": report_data.get("estimate_tasks", []), "EOD Reminder Tasks": report_data.get("reminder_tasks", []), "EOD Shipping Tasks": report_data.get("shipping_tasks", []), "EOD AdHoc Shipped": report_data.get("adhoc_shipped_today", []) }
        else:
            return BytesIO().getvalue()

        for sheet_name_key, data_list in sheets_data.items():
            df_report_sheet = pd.DataFrame(data_list)
            if not df_report_sheet.empty:
                df_report_sheet.to_excel(writer, sheet_name=sheet_name_key, startrow=2, index=False, header=False)
                worksheet = writer.sheets[sheet_name_key]
                worksheet.merge_range(0, 0, 0, len(df_report_sheet.columns)-1 if len(df_report_sheet.columns)>0 else 0, f"{sheet_name_key} - Report Date: {report_date_for_title}", title_format)
                worksheet.set_row(0, 30)
                for col_num, value in enumerate(df_report_sheet.columns.values): worksheet.write(2, col_num, value, header_format)
                for row_num in range(3, len(df_report_sheet) + 3):
                    for col_num in range(len(df_report_sheet.columns)):
                        worksheet.write(row_num, col_num, df_report_sheet.iloc[row_num-3, col_num], cell_format)
                for i, col_name_iter in enumerate(df_report_sheet.columns):
                    column_width = max(df_report_sheet[col_name_iter].astype(str).map(len).max(), len(str(col_name_iter))) + 2
                    worksheet.set_column(i, i, column_width)
            else:
                worksheet = writer.book.add_worksheet(sheet_name_key)
                worksheet.merge_range(0, 0, 0, 2, f"{sheet_name_key} - Report Date: {report_date_for_title}", title_format)
                worksheet.write(2,0, "No items for this category.", cell_format)
    return output.getvalue()
@st.cache_data(ttl=300)

def load_price_library_df():
    '''Loads the entire price library from Google Sheets into a DataFrame for display and editing.'''
    client = connect_to_google_sheet()
    if not client:
        return pd.DataFrame()
    try:
        spreadsheet = client.open(GSHEET_NAME)
        worksheet = spreadsheet.worksheet(PRICE_LIBRARY_SHEET_NAME)
        data = worksheet.get_all_values()
        if not data or len(data) < 1:
            return pd.DataFrame()
        
        headers = data.pop(0)
        df = pd.DataFrame(data, columns=headers)
        
        if 'Amount Including Tax' in df.columns:
            df['Amount Including Tax'] = pd.to_numeric(df['Amount Including Tax'], errors='coerce').fillna(0.0)
        if 'Usage Count' in df.columns:
            df['Usage Count'] = pd.to_numeric(df['Usage Count'], errors='coerce').fillna(0).astype(int)
            
        return df
    except gspread.exceptions.WorksheetNotFound:
        st.info("Price Library sheet not found. It can be created by generating an estimate.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error loading Price Library as a table: {e}")
        return pd.DataFrame()

def save_price_library_df(df):
    '''Saves an entire DataFrame to the Price Library sheet, overwriting all existing data.'''
    client = connect_to_google_sheet()
    if not client:
        return False
    try:
        spreadsheet = client.open(GSHEET_NAME)
        worksheet = spreadsheet.worksheet(PRICE_LIBRARY_SHEET_NAME)
        
        df_to_save = df.astype(str)
        
        worksheet.clear()
        worksheet.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist(), value_input_option='USER_ENTERED')
        
        return True
    except Exception as e:
        st.error(f"Failed to save Price Library: {e}")
        return False
    

def get_revision_rma(original_rma):
    """Generates the next revision number for an RMA (e.g., RMA123 -> RMA123-R1)."""
    base_rma = str(original_rma).split('-R')[0] if original_rma else "UnknownRMA"
    
    parts = str(original_rma).split('-R')
    rev_num = 1
    if len(parts) > 1 and parts[1].isdigit():
        rev_num = int(parts[1]) + 1
        
    return f"{base_rma}-R{rev_num}"

def load_estimate_for_revision(rma_to_find):
    """
    Finds the latest entry for an RMA, and attempts to load its original parts list
    from the archive before falling back to JSON.
    """
    client = connect_to_google_sheet()
    if not client:
        return None
    try:
        worksheet = client.open(GSHEET_NAME).worksheet(ESTIMATE_SHEET_NAME)
        all_records = worksheet.get_all_records()
        
        search_term = str(rma_to_find).strip()
        
        matching_records = []
        for rec in all_records:
            rma_from_sheet_raw = rec.get('RMA', '')
            rma_from_sheet_str = str(rma_from_sheet_raw).strip()
            
            if search_term.isdigit() and rma_from_sheet_str.isdigit():
                if int(search_term) == int(rma_from_sheet_str):
                    matching_records.append(rec)
                    continue 
            
            base_rma_from_sheet = rma_from_sheet_str.split('-R')[0]
            if rma_from_sheet_str.startswith(search_term) or base_rma_from_sheet.startswith(search_term):
                matching_records.append(rec)

        if not matching_records:
            return None
            
        latest_record = matching_records[-1]

        # --- MODIFIED LOGIC: Check for both original and standardized (6-digit) filenames ---
        rma_from_record = str(latest_record.get('RMA', '')).strip()

        # Path for the filename exactly as it is in the sheet (e.g., "12345.xlsx")
        original_path = os.path.join(SOURCE_PARTS_ARCHIVE_DIR, f"{rma_from_record}.xlsx")

        # Path for a standardized 6-digit filename (e.g., "012345.xlsx")
        standardized_path = None
        if rma_from_record.isdigit():
            # Only create a standardized path if the RMA is purely numeric
            standardized_filename = rma_from_record.zfill(5) + ".xlsx"
            standardized_path = os.path.join(SOURCE_PARTS_ARCHIVE_DIR, standardized_filename)

        # --- DEBUGGING MESSAGES ---
        st.info(f"Checking for original path: '{os.path.abspath(original_path)}'")
        if standardized_path:
            st.info(f"Checking for standardized path: '{os.path.abspath(standardized_path)}'")

        # --- Check which path exists, preferring the standardized one ---
        path_to_load = None
        if standardized_path and os.path.exists(standardized_path):
            path_to_load = standardized_path
        elif os.path.exists(original_path):
            path_to_load = original_path

        if path_to_load:
            st.success(f"Found archived parts file for {rma_from_record}. Loading automatically from '{os.path.basename(path_to_load)}'.")
            parts_df = pd.read_excel(path_to_load)
            latest_record['parts_df'] = parts_df
            latest_record['is_legacy'] = False
            return latest_record
        # --- END OF MODIFIED LOGIC ---

        # Fallback to the original method (checking for Parts JSON)
        parts_json = latest_record.get('Parts JSON', '')
        is_legacy = not bool(parts_json) 

        parts_df = pd.read_json(parts_json, orient='records') if not is_legacy else pd.DataFrame()
        
        latest_record['parts_df'] = parts_df
        latest_record['is_legacy'] = is_legacy
        return latest_record

    except Exception as e:
        st.error(f"Error loading estimate for revision: {e}")
        return None
    
def process_historical_usage(master_df):
    """
    Processes a master dataframe of all parts from historical uploads,
    and updates the Price Library with quarterly usage counts.
    """
    if 'Planned Delivery Date' not in master_df.columns:
        st.error("The uploaded files are missing the 'Planned Delivery Date' column.")
        return False

    master_df['Estimate Date'] = pd.to_datetime(master_df['Planned Delivery Date'], errors='coerce')
    master_df.dropna(subset=['Estimate Date'], inplace=True)

    master_df['Quarter'] = master_df['Estimate Date'].dt.to_period('Q').astype(str).str.replace('Q', 'Q-')

    quarterly_counts = master_df.groupby(['No.', 'Quarter']).size().unstack(fill_value=0)
    latest_part_info = master_df.drop_duplicates(subset=['No.'], keep='last').set_index('No.')
    latest_part_info = latest_part_info[['Description', 'Amount Including Tax']]

    try:
        library_df = load_price_library_df()
        if not library_df.empty:
            if 'Usage Count' in library_df.columns:
                library_df.rename(columns={'Usage Count': 'Total Usage'}, inplace=True)
            library_df.set_index('No.', inplace=True)
        else:
            library_df = pd.DataFrame(columns=['No.', 'Description', 'Amount Including Tax', 'Total Usage']).set_index('No.')
    except Exception as e:
        st.error(f"Could not load existing price library: {e}")
        return False

    for part_no, row in quarterly_counts.iterrows():
        for quarter, count in row.items():
            if quarter not in library_df.columns:
                library_df[quarter] = 0
            if part_no not in library_df.index:
                library_df.loc[part_no] = 0
            library_df.loc[part_no, quarter] += count

    library_df.update(latest_part_info)
    
    quarter_cols = [col for col in library_df.columns if 'Q-' in str(col)]
    library_df['Total Usage'] = library_df[quarter_cols].sum(axis=1)

    library_df.reset_index(inplace=True)
    final_cols = ['No.', 'Description', 'Amount Including Tax', 'Total Usage'] + sorted(quarter_cols)
    for col in final_cols:
        if col not in library_df.columns:
            library_df[col] = 0
            
    library_df = library_df[final_cols]
    
    return save_price_library_df(library_df)

@st.cache_data(ttl=600)
def load_customer_list():
    """Loads the customer list from Google Sheets into a dictionary."""
    CUSTOMER_LIST_SHEET_NAME = "Customer List"
    client = connect_to_google_sheet()
    if not client:
        return {}
    try:
        spreadsheet = client.open(GSHEET_NAME)
        worksheet = spreadsheet.worksheet(CUSTOMER_LIST_SHEET_NAME)
        records = worksheet.get_all_records()
        customer_map = {
            str(rec.get('Customer Number')): rec.get('Customer Name')
            for rec in records if 'Customer Number' in rec and 'Customer Name' in rec
        }
        return customer_map
    except gspread.exceptions.WorksheetNotFound:
        st.warning(f"'{CUSTOMER_LIST_SHEET_NAME}' sheet not found. Autofill for customer name is disabled.")
        return {}
    except Exception as e:
        st.error(f"Error loading Customer List: {e}")
        return {}
    

@st.cache_data(ttl=1800)
def load_zone_ranges():
    """Loads the ZIP code ranges and corresponding zones from Google Sheets."""
    client = connect_to_google_sheet()
    if not client: return pd.DataFrame()
    try:
        worksheet = client.open(GSHEET_NAME).worksheet("Shipping Zones")
        data = worksheet.get_all_records()
        if not data: return pd.DataFrame()

        df = pd.DataFrame(data)
        df['Start ZIP'] = pd.to_numeric(df['Start ZIP'], errors='coerce')
        df['End ZIP'] = pd.to_numeric(df['End ZIP'], errors='coerce')
        df.dropna(subset=['Start ZIP', 'End ZIP'], inplace=True)
        df['Start ZIP'] = df['Start ZIP'].astype(int)
        df['End ZIP'] = df['End ZIP'].astype(int)
        return df
    except gspread.exceptions.WorksheetNotFound:
        st.warning("'Shipping Zones' sheet not found. Shipping automation is disabled.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error loading Shipping Zones: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=1800)
def load_shipping_prices():
    """Loads the shipping prices and sets the Zone as the index."""
    client = connect_to_google_sheet()
    if not client: return pd.DataFrame()
    try:
        worksheet = client.open(GSHEET_NAME).worksheet("Shipping Prices")
        data = worksheet.get_all_records()
        if not data: return pd.DataFrame()
        
        df = pd.DataFrame(data)

        df.columns = [str(col).strip() for col in df.columns]

        if 'Zone' in df.columns:
            df['Zone'] = df['Zone'].astype(str)
            df.set_index('Zone', inplace=True)
        return df
    except gspread.exceptions.WorksheetNotFound:
        st.warning("'Shipping Prices' sheet not found. Shipping automation is disabled.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error loading Shipping Prices: {e}")
        return pd.DataFrame()
