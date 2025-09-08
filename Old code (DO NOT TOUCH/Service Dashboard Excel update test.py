import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta 
import gspread 
from oauth2client.service_account import ServiceAccountCredentials 
from io import BytesIO 
import urllib.parse 
import json # For storing list of dicts as string in GSheet
# import xlsxwriter # Not directly imported if using pandas.ExcelWriter engine, but good to have installed

# --- Page Configuration ---
st.set_page_config(
    page_title="Service Data Dashboard", 
    page_icon="🚚", 
    layout="wide",
)

# --- Constants for Google Sheets ---
GSHEET_NAME = "Estimate form"
WORKSHEET_INDEX = 1 # Main data sheet
CREDS_FILE = "Credentials.json" 
ARCHIVE_SHEET_NAME = "DailyReportArchive" 
EOD_SUMMARY_ARCHIVE_SHEET_NAME = "EODSummaryArchive" # New sheet for EOD summary archive

ARCHIVE_SHEET_HEADERS = ["Report Date", "Needs Estimate Creation", "Needs Shipping", "Needs Reminder"]
EOD_ARCHIVE_SHEET_HEADERS = ["Report Date", "Estimate Task Summary", "Reminder Task Summary", "Shipping Task Summary", "AdHoc Shipped Today"]


EXPECTED_COLUMN_ORDER = [
    "RMA", "SPC Code", "Part Number", "S/N", "Description", 
    "Fault Comments", "Resolution Comments", "Sender", 
    "Estimate Complete Time", "Estimate Complete", 
    "Estimate Approved", "Estimate Approved Time",
    "Estimate Sent To Email", "Estimate Sent Time", 
    "Reminder Completed", "Reminder Completed Time", "Reminder Contact Method", 
    "QA Approved", "QA Approved Time",
    "Shipped", "Shipped Time" 
]
ALL_STATUS_COLUMNS = ["Estimate Complete", "Estimate Approved", "Reminder Completed", "QA Approved", "Shipped"]
ALL_TIME_COLUMNS = [col for col in EXPECTED_COLUMN_ORDER if "Time" in col]


# --- Constants for Business Central Link ---
BC_BASE_URL = "https://businesscentral.dynamics.com/7bcfb5b0-27a1-4e18-99d8-ca66570addd8/Production"
BC_COMPANY = "PROD"
BC_PAGE_ID = "70001" 
BC_RMA_FIELD_NAME = "No." 
BC_LINK_COL_NAME = "View in BC" 

# --- Helper Functions ---
@st.cache_data(ttl=300) 
def load_data_from_google_sheet(
    sheet_name=GSHEET_NAME, 
    worksheet_index=WORKSHEET_INDEX, 
    creds_file=CREDS_FILE
):
    """Loads data from the specified Google Sheet."""
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open(sheet_name)
        worksheet = spreadsheet.get_worksheet(worksheet_index)
        
        all_values = worksheet.get_all_values()
        
        if not all_values:
            # st.warning(f"No data (not even headers) found in Google Sheet '{sheet_name}', worksheet index {worksheet_index}.") # Can be noisy
            return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) 
            
        headers_from_sheet = all_values[0]
        data_rows = all_values[1:]
        
        temp_df = pd.DataFrame(data_rows, columns=headers_from_sheet)
        df = pd.DataFrame(columns=EXPECTED_COLUMN_ORDER)

        for col in EXPECTED_COLUMN_ORDER:
            if col in temp_df.columns:
                df[col] = temp_df[col] 
            else: 
                 # st.warning(f"Expected column '{col}' not found in Google Sheet. Initializing as empty/default.") # Less verbose
                 if "Time" in col: df[col] = pd.NaT
                 elif col in ALL_STATUS_COLUMNS: df[col] = "No"
                 elif col == "Estimate Sent To Email" or col == "Reminder Contact Method": df[col] = "N/A" 
                 else: df[col] = "N/A" 
        
        df = df[EXPECTED_COLUMN_ORDER] 

        string_cols_to_process = ['RMA', 'S/N', 'Part Number', 'SPC Code', 
                                   'Description', 'Fault Comments', 'Resolution Comments', 
                                   'Sender', 'Estimate Sent To Email', 'Reminder Contact Method'] + ALL_STATUS_COLUMNS
        for col in string_cols_to_process:
            if col in df.columns:
                df[col] = df[col].astype(str) 
                if col in ALL_STATUS_COLUMNS:
                     df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'No') 
                elif col == "Estimate Sent To Email" or col == "Reminder Contact Method":
                     df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'N/A')
                else: 
                     df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'N/A')

        for col in ALL_TIME_COLUMNS:
            if col in df.columns:
                df[col] = df[col].replace('N/A', None) 
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        return df
    except Exception as e:
        st.error(f"An error occurred while loading data from Google Sheets: {type(e).__name__} - {e}")
    return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) 


def find_row_in_gsheet(worksheet, rma_to_find, sn_to_find, headers):
    """
    Finds a row in the worksheet based on RMA and S/N.
    Performs case-insensitive and whitespace-stripped comparison.
    """
    try:
        rma_col_idx = headers.index("RMA") 
        sn_col_idx = headers.index("S/N") 
    except ValueError:
        # This error is critical for updates, so it should be visible.
        # st.error("Critical error: RMA or S/N column header not found in the Google Sheet. Cannot perform updates.")
        return -1 

    all_data_values = worksheet.get_all_values() 
    for i, row_values in enumerate(all_data_values[1:], start=2): 
        rma_val_from_sheet = row_values[rma_col_idx] if len(row_values) > rma_col_idx else None
        sn_val_from_sheet = row_values[sn_col_idx] if len(row_values) > sn_col_idx else None
        
        if rma_val_from_sheet is not None and sn_val_from_sheet is not None:
            rma_to_find_str = str(rma_to_find).strip().lower()
            sn_to_find_str = str(sn_to_find).strip().lower()
            sheet_rma_str = str(rma_val_from_sheet).strip().lower()
            sheet_sn_str = str(sn_val_from_sheet).strip().lower()
            if sheet_rma_str == rma_to_find_str and sheet_sn_str == sn_to_find_str:
                return i 
    return -1 

def update_gsheet_cells(worksheet, updates_list):
    try:
        worksheet.batch_update(updates_list)
        return True
    except Exception as e: st.error(f"An error occurred during Google Sheet batch update: {e}"); return False

def gsheet_update_wrapper(update_function, *args):
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open(GSHEET_NAME)
        worksheet = spreadsheet.get_worksheet(WORKSHEET_INDEX) 
        headers = worksheet.row_values(1)
        if not headers: st.error("Could not read headers from main data sheet. Update failed."); return False
        return update_function(worksheet, headers, *args)
    except Exception as e: st.error(f"General error during Google Sheet operation: {type(e).__name__} - {e}"); return False

def _update_estimate_sent_in_sheet(worksheet, headers, rma, sn, sent_to_email, sent_date_obj):
    sent_time_col_name = "Estimate Sent Time"; sent_email_col_name = "Estimate Sent To Email"
    try:
        sent_time_col_idx = headers.index(sent_time_col_name) + 1
        sent_email_col_idx = headers.index(sent_email_col_name) + 1
    except ValueError: st.error(f"'{sent_time_col_name}' or '{sent_email_col_name}' not in sheet headers."); return False
    row_to_update = find_row_in_gsheet(worksheet, rma, sn, headers)
    if row_to_update != -1:
        sent_time_str = datetime.combine(sent_date_obj, datetime.now().time()).strftime("%Y-%m-%d %H:%M:%S")
        updates = [
            {'range': gspread.utils.rowcol_to_a1(row_to_update, sent_email_col_idx), 'values': [[sent_to_email]]},
            {'range': gspread.utils.rowcol_to_a1(row_to_update, sent_time_col_idx), 'values': [[sent_time_str]]} ]
        if update_gsheet_cells(worksheet, updates):
            st.success(f"Estimate for RMA {rma}, S/N {sn} marked as sent to {sent_to_email} on {sent_date_obj.strftime('%Y-%m-%d')}."); return True
    else: st.error(f"Record for RMA {rma}, S/N {sn} not found for estimate sent update.")
    return False

def _update_reminder_in_sheet(worksheet, headers, rma, sn, reminder_date_obj, contact_method): 
    reminder_status_col_name = "Reminder Completed"; reminder_time_col_name = "Reminder Completed Time"
    reminder_method_col_name = "Reminder Contact Method" 
    try:
        reminder_status_col_idx = headers.index(reminder_status_col_name) + 1
        reminder_time_col_idx = headers.index(reminder_time_col_name) + 1
        reminder_method_col_idx = headers.index(reminder_method_col_name) + 1 
    except ValueError: st.error(f"One of the reminder columns not in sheet headers."); return False
    
    row_to_update = find_row_in_gsheet(worksheet, rma, sn, headers)
    if row_to_update != -1:
        reminder_time_str = datetime.combine(reminder_date_obj, datetime.now().time()).strftime("%Y-%m-%d %H:%M:%S")
        updates = [
            {'range': gspread.utils.rowcol_to_a1(row_to_update, reminder_status_col_idx), 'values': [["Yes"]]},
            {'range': gspread.utils.rowcol_to_a1(row_to_update, reminder_time_col_idx), 'values': [[reminder_time_str]]},
            {'range': gspread.utils.rowcol_to_a1(row_to_update, reminder_method_col_idx), 'values': [[contact_method]]} 
        ]
        if update_gsheet_cells(worksheet, updates):
            st.success(f"Reminder for RMA {rma}, S/N {sn} (via {contact_method}) marked as completed on {reminder_date_obj.strftime('%Y-%m-%d')}."); return True
    else: st.error(f"Record for RMA {rma}, S/N {sn} not found for reminder update.")
    return False

def _update_shipped_in_sheet(worksheet, headers, rma, sn, shipped_date_obj):
    shipped_status_col_name = "Shipped"; shipped_time_col_name = "Shipped Time"
    try:
        shipped_status_col_idx = headers.index(shipped_status_col_name) + 1
        shipped_time_col_idx = headers.index(shipped_time_col_name) + 1
    except ValueError: st.error(f"'{shipped_status_col_name}' or '{shipped_time_col_name}' not in sheet headers."); return False
    row_to_update = find_row_in_gsheet(worksheet, rma, sn, headers)
    if row_to_update != -1:
        shipped_time_str = datetime.combine(shipped_date_obj, datetime.now().time()).strftime("%Y-%m-%d %H:%M:%S")
        updates = [
            {'range': gspread.utils.rowcol_to_a1(row_to_update, shipped_status_col_idx), 'values': [["Yes"]]},
            {'range': gspread.utils.rowcol_to_a1(row_to_update, shipped_time_col_idx), 'values': [[shipped_time_str]]} ]
        if update_gsheet_cells(worksheet, updates):
            st.success(f"Successfully marked RMA {rma}, S/N {sn} as shipped on {shipped_date_obj.strftime('%Y-%m-%d')}."); return True
    else: st.error(f"Record with RMA {rma} and S/N {sn} not found in Google Sheet. Update failed.")
    return False

def update_estimate_sent_details_in_gsheet(rma, sn, sent_to_email, sent_date_obj):
    return gsheet_update_wrapper(_update_estimate_sent_in_sheet, rma, sn, sent_to_email, sent_date_obj)
def update_reminder_details_in_gsheet(rma, sn, reminder_date_obj, contact_method): 
    return gsheet_update_wrapper(_update_reminder_in_sheet, rma, sn, reminder_date_obj, contact_method)
def update_shipped_status_in_gsheet(rma, sn, shipped_date_obj): 
    return gsheet_update_wrapper(_update_shipped_in_sheet, rma, sn, shipped_date_obj)

def display_kpis(df):
    if df.empty: return
    total_records = len(df)
    kpi_cols = { "Est. Complete": 'Estimate Complete', "Est. Approved": 'Estimate Approved',
        "Est. Sent": 'Estimate Sent To Email', "Reminders Done": 'Reminder Completed', 
        "QA Approved": 'QA Approved', "Units Shipped": 'Shipped' }
    kpi_values = {"Total Records": total_records}
    for label, col_name in kpi_cols.items():
        if col_name in df.columns:
            if col_name == 'Estimate Sent To Email': kpi_values[label] = df[df[col_name].astype(str).str.lower() != 'n/a'].shape[0]
            else: kpi_values[label] = df[df[col_name].astype(str).str.lower() == 'yes'].shape[0]
        else: kpi_values[label] = 0
    cols = st.columns(len(kpi_values))
    for i, (label, value) in enumerate(kpi_values.items()): cols[i].metric(label, value)

def identify_overdue_estimates(df, days_threshold=3):
    required_cols = ['Estimate Complete Time', 'Estimate Complete', 'Estimate Approved', 'RMA', 'S/N', 'SPC Code', 
                     'Estimate Sent To Email', 'Estimate Sent Time', 
                     'Reminder Completed', 'Reminder Completed Time', 'Reminder Contact Method'] 
    if df.empty or not all(col in df.columns for col in required_cols): return pd.DataFrame()
    df_copy = df.copy(); df_copy['Estimate Complete Time'] = pd.to_datetime(df_copy['Estimate Complete Time'], errors='coerce')
    df_copy['Estimate Sent Time'] = pd.to_datetime(df_copy['Estimate Sent Time'], errors='coerce')
    df_copy['Reminder Completed Time'] = pd.to_datetime(df_copy['Reminder Completed Time'], errors='coerce')
    now = datetime.now(); overdue_items = []
    for _, row in df_copy.iterrows():
        if str(row.get('Estimate Complete', 'N/A')).lower() == 'yes' and \
           str(row.get('Estimate Approved', 'N/A')).lower() in ['no', 'n/a'] and \
           pd.notna(row['Estimate Complete Time']):
            if (now - row['Estimate Complete Time']).days > days_threshold:
                rma_value = str(row.get('RMA', 'N/A'))
                bc_url = f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(rma_value)}%27" if rma_value not in ['N/A', ''] else None
                est_sent_time_str = row['Estimate Sent Time'].strftime('%Y-%m-%d %H:%M') if pd.notna(row['Estimate Sent Time']) else 'N/A'
                reminder_time_str = row['Reminder Completed Time'].strftime('%Y-%m-%d %H:%M') if pd.notna(row['Reminder Completed Time']) else 'N/A'
                overdue_items.append({
                    'RMA': rma_value, 'S/N': row.get('S/N', 'N/A'), 'SPC Code': row.get('SPC Code', 'N/A'), 
                    'Estimate Complete Time': row['Estimate Complete Time'].strftime('%Y-%m-%d'),
                    'Days Pending Approval': (now - row['Estimate Complete Time']).days,
                    'Estimate Sent To Email': row.get('Estimate Sent To Email', 'N/A'), 
                    'Estimate Sent Time': est_sent_time_str,  
                    'Reminder Completed': row.get('Reminder Completed', 'N/A'), 
                    'Reminder Completed Time': reminder_time_str,
                    'Reminder Contact Method': row.get('Reminder Contact Method', 'N/A'), 
                    BC_LINK_COL_NAME: bc_url  })
    return pd.DataFrame(overdue_items)

def identify_overdue_for_shipping(df, days_threshold=1):
    required_cols = ['QA Approved Time', 'Estimate Complete', 'Estimate Approved', 'QA Approved', 'Shipped', 'RMA', 'S/N', 'SPC Code'] 
    if df.empty or not all(col in df.columns for col in required_cols): return pd.DataFrame()
    df_copy = df.copy(); df_copy['QA Approved Time'] = pd.to_datetime(df_copy['QA Approved Time'], errors='coerce')
    now = datetime.now(); overdue_shipping_items = []
    for _, row in df_copy.iterrows():
        if str(row.get('Estimate Complete', 'N/A')).lower() == 'yes' and \
           str(row.get('Estimate Approved', 'N/A')).lower() == 'yes' and \
           str(row.get('QA Approved', 'N/A')).lower() == 'yes' and \
           str(row.get('Shipped', 'N/A')).lower() in ['no', 'n/a'] and \
           pd.notna(row['QA Approved Time']):
            if (now - row['QA Approved Time']).days > days_threshold:
                rma_value = str(row.get('RMA', 'N/A'))
                bc_url = f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(rma_value)}%27" if rma_value not in ['N/A', ''] else None
                overdue_shipping_items.append({ 
                    'RMA': rma_value, 'S/N': row.get('S/N', 'N/A'), 'SPC Code': row.get('SPC Code', 'N/A'), 
                    'QA Approved Time': row['QA Approved Time'].strftime('%Y-%m-%d'),
                    'Days Pending Shipping': (now - row['QA Approved Time']).days, BC_LINK_COL_NAME: bc_url })
    return pd.DataFrame(overdue_shipping_items)

def identify_overdue_reminders(df, days_threshold=2):
    required_cols = ['Estimate Sent Time', 'Estimate Sent To Email', 'Reminder Completed', 'RMA', 'S/N', 'SPC Code', 'Reminder Contact Method'] 
    if df.empty or not all(col in df.columns for col in required_cols): return pd.DataFrame()
    df_copy = df.copy(); df_copy['Estimate Sent Time'] = pd.to_datetime(df_copy['Estimate Sent Time'], errors='coerce')
    now = datetime.now(); overdue_reminder_items = []
    for _, row in df_copy.iterrows():
        is_estimate_sent = str(row.get('Estimate Sent To Email', 'N/A')).lower() != 'n/a'
        is_reminder_not_done = str(row.get('Reminder Completed', 'N/A')).lower() in ['no', 'n/a']
        estimate_sent_time = row['Estimate Sent Time']; rma_value = str(row.get('RMA', 'N/A'))
        if is_estimate_sent and is_reminder_not_done and pd.notna(estimate_sent_time):
            days_passed_reminder = (now - estimate_sent_time).days
            if days_passed_reminder > days_threshold:
                bc_url = f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(rma_value)}%27" if rma_value not in ['N/A', ''] else None
                overdue_reminder_items.append({
                    'RMA': rma_value, 'S/N': row.get('S/N', 'N/A'), 'SPC Code': row.get('SPC Code', 'N/A'), 
                    'Estimate Sent To Email': row.get('Estimate Sent To Email', 'N/A'),
                    'Estimate Sent Time': estimate_sent_time.strftime('%Y-%m-%d') if pd.notna(estimate_sent_time) else 'N/A',
                    'Days Pending Reminder': days_passed_reminder,
                    'Reminder Contact Method': row.get('Reminder Contact Method', 'N/A'), 
                    BC_LINK_COL_NAME: bc_url  })
    return pd.DataFrame(overdue_reminder_items)

# --- Daily Status Report Functions (Modified for GSheet Archive) ---
@st.cache_data(ttl=60) 
def get_archived_reports_from_gsheet(archive_sheet_name, expected_headers):
    """Loads all archived reports from the specified Google Sheet archive tab."""
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open(GSHEET_NAME)
        try:
            archive_ws = spreadsheet.worksheet(archive_sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"Archive sheet '{archive_sheet_name}' not found. Please create it with headers: {', '.join(expected_headers)}.")
            return []
        
        records = archive_ws.get_all_records() 
        archived_reports = []
        
        if archive_sheet_name == ARCHIVE_SHEET_NAME: # Daily Status Report Archive
            for rec in records:
                try:
                    needs_est_str = rec.get('Needs Estimate Creation', '[]')
                    needs_ship_str = rec.get('Needs Shipping', '[]')
                    needs_reminder_str = rec.get('Needs Reminder', '[]') 
                    needs_est_list = json.loads(needs_est_str) if needs_est_str and needs_est_str.strip() else []
                    needs_ship_list = json.loads(needs_ship_str) if needs_ship_str and needs_ship_str.strip() else []
                    needs_reminder_list = json.loads(needs_reminder_str) if needs_reminder_str and needs_reminder_str.strip() else [] 
                    archived_reports.append({
                        "date": rec.get('Report Date'),
                        "needs_estimate_creation": needs_est_list,
                        "needs_shipping": needs_ship_list,
                        "needs_reminder": needs_reminder_list })
                except Exception: pass # Skip malformed rows
        elif archive_sheet_name == EOD_SUMMARY_ARCHIVE_SHEET_NAME:
            for rec in records:
                try:
                    archived_reports.append({
                        "date": rec.get('Report Date'),
                        "estimate_tasks": json.loads(rec.get('Estimate Task Summary', '[]')),
                        "reminder_tasks": json.loads(rec.get('Reminder Task Summary', '[]')),
                        "shipping_tasks": json.loads(rec.get('Shipping Task Summary', '[]')),
                        "adhoc_shipped_today": json.loads(rec.get('AdHoc Shipped Today', '[]'))
                    })
                except Exception: pass # Skip malformed rows
        
        archived_reports.sort(key=lambda r: r.get('date', ''), reverse=True)
        return archived_reports
    except Exception as e:
        st.error(f"Error loading archived reports from '{archive_sheet_name}': {type(e).__name__} - {e}")
        return []

def get_last_report_date_from_archive(archived_reports):
    if not archived_reports: return date.today() - timedelta(days=1) 
    try:
        latest_date_str = archived_reports[0]['date']
        return datetime.strptime(latest_date_str, "%Y-%m-%d").date()
    except: return date.today() - timedelta(days=1) 

def save_report_to_gsheet_archive(report_data, archive_sheet_name_to_save, archive_headers_to_check):
    """Saves a single daily report to the specified Google Sheet archive."""
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open(GSHEET_NAME)
        try:
            archive_ws = spreadsheet.worksheet(archive_sheet_name_to_save)
            if archive_ws.row_count == 0 or archive_ws.row_values(1) != archive_headers_to_check:
                archive_ws.clear(); archive_ws.append_row(archive_headers_to_check)
                st.info(f"Archive sheet '{archive_sheet_name_to_save}' headers reset/created.")
        except gspread.exceptions.WorksheetNotFound:
            st.info(f"Archive sheet '{archive_sheet_name_to_save}' not found. Creating it with headers: {', '.join(archive_headers_to_check)}.")
            archive_ws = spreadsheet.add_worksheet(title=archive_sheet_name_to_save, rows="1", cols=str(len(archive_headers_to_check)))
            archive_ws.append_row(archive_headers_to_check)
            
        existing_dates = archive_ws.col_values(1)[1:] 
        # For EOD reports, allow overwriting if one for the same date exists.
        # For Daily reports, we skip if it exists.
        if report_data['date'] in existing_dates and archive_sheet_name_to_save == ARCHIVE_SHEET_NAME: 
            return False 
        
        row_to_append = [report_data['date']]
        if archive_sheet_name_to_save == ARCHIVE_SHEET_NAME:
            row_to_append.extend([
                json.dumps(report_data['needs_estimate_creation']),
                json.dumps(report_data['needs_shipping']),
                json.dumps(report_data['needs_reminder'])
            ])
        elif archive_sheet_name_to_save == EOD_SUMMARY_ARCHIVE_SHEET_NAME:
             row_to_append.extend([
                json.dumps(report_data['estimate_tasks']),
                json.dumps(report_data['reminder_tasks']),
                json.dumps(report_data['shipping_tasks']),
                json.dumps(report_data.get('adhoc_shipped_today', [])) 
            ])
        
        # If EOD report for the date already exists, find and update row instead of appending
        if archive_sheet_name_to_save == EOD_SUMMARY_ARCHIVE_SHEET_NAME and report_data['date'] in existing_dates:
            try:
                cell = archive_ws.find(report_data['date']) # Find cell with the date
                archive_ws.update(f'A{cell.row}', [row_to_append]) # Update the entire row
                st.info(f"EOD Summary for {report_data['date']} updated in archive.")
            except gspread.exceptions.CellNotFound: # Should not happen if date is in existing_dates
                 archive_ws.append_row(row_to_append) 
        else:
            archive_ws.append_row(row_to_append) 
        
        # Clear specific cache based on which archive was updated
        if archive_sheet_name_to_save == ARCHIVE_SHEET_NAME:
            get_archived_reports_from_gsheet.clear(archive_sheet_name=ARCHIVE_SHEET_NAME, expected_headers=ARCHIVE_SHEET_HEADERS)
        elif archive_sheet_name_to_save == EOD_SUMMARY_ARCHIVE_SHEET_NAME:
            get_archived_reports_from_gsheet.clear(archive_sheet_name=EOD_SUMMARY_ARCHIVE_SHEET_NAME, expected_headers=EOD_ARCHIVE_SHEET_HEADERS)
        return True
    except Exception as e: st.error(f"Error saving report to '{archive_sheet_name_to_save}': {type(e).__name__} - {e}"); return False

def generate_single_day_report_content(df, report_date_obj):
    report_content = { "date": report_date_obj.strftime("%Y-%m-%d"), 
                      "needs_shipping": [], 
                      "needs_estimate_creation": [],
                      "needs_reminder": [] } 
    
    # Needs Shipping
    shipping_df = df[
        (df['Estimate Complete'].astype(str).str.lower() == 'yes') & 
        (df['Estimate Approved'].astype(str).str.lower() == 'yes') & 
        (df['QA Approved'].astype(str).str.lower() == 'yes') &
        (df['Shipped'].astype(str).str.lower().isin(['no', 'n/a'])) &
        (pd.to_datetime(df['QA Approved Time'], errors='coerce').dt.date == report_date_obj) ] 
    for _, row in shipping_df.iterrows():
        report_content["needs_shipping"].append({'RMA': str(row['RMA']), 'S/N': str(row['S/N']), 'SPC Code': str(row.get('SPC Code', 'N/A'))})
    
    # Needs Estimate Creation
    day_prior_to_report = report_date_obj - timedelta(days=1)
    estimate_df = df[
        (df['Estimate Complete'].astype(str).str.lower() == 'yes') &
        (df['Estimate Sent To Email'].astype(str).str.lower() == 'n/a') & 
        (pd.to_datetime(df['Estimate Complete Time'], errors='coerce').dt.date == day_prior_to_report) ] 
    for _, row in estimate_df.iterrows():
        report_content["needs_estimate_creation"].append({
            'RMA': str(row['RMA']), 'S/N': str(row['S/N']), 'SPC Code': str(row.get('SPC Code', 'N/A')),
            'Est. Complete Date': day_prior_to_report.strftime('%Y-%m-%d') })
            
    # Needs Reminder (Estimate Sent 2 days before report_date_obj, Reminder Not Completed)
    estimate_sent_target_date = report_date_obj - timedelta(days=2)
    reminder_df = df[
        (df['Estimate Sent To Email'].astype(str).str.lower() != 'n/a') &
        (df['Reminder Completed'].astype(str).str.lower().isin(['no', 'n/a'])) &
        (pd.to_datetime(df['Estimate Sent Time'], errors='coerce').dt.date == estimate_sent_target_date)
    ]
    for _, row in reminder_df.iterrows():
        report_content["needs_reminder"].append({
            'RMA': str(row['RMA']), 
            'S/N': str(row['S/N']),
            'SPC Code': str(row.get('SPC Code', 'N/A')),
            'Estimate Sent To Email': str(row['Estimate Sent To Email']),
            'Estimate Sent Time': pd.to_datetime(row['Estimate Sent Time']).strftime('%Y-%m-%d') if pd.notna(row['Estimate Sent Time']) else 'N/A'
        })
    return report_content

def create_excel_report_bytes(report_data, report_type="Daily"):
    """Creates an Excel file in bytes from the structured report data with improved formatting."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer: 
        workbook = writer.book
        header_format = workbook.add_format({'bold': True, 'text_wrap': False, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
        cell_format = workbook.add_format({'border': 1})
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})

        if report_type == "Daily" or report_type == "Custom":
            sheets_data = {
                "Needs Estimate Creation": report_data.get("needs_estimate_creation", []),
                "Needs Reminder": report_data.get("needs_reminder", []), 
                "Needs Shipping": report_data.get("needs_shipping", [])
            }
        elif report_type == "EOD":
            sheets_data = {
                "EOD Estimate Tasks": report_data.get("estimate_tasks", []),
                "EOD Reminder Tasks": report_data.get("reminder_tasks", []), 
                "EOD Shipping Tasks": report_data.get("shipping_tasks", []),
                "EOD AdHoc Shipped": report_data.get("adhoc_shipped_today", [])
            }
        else:
            return output.getvalue() # Should not happen

        for sheet_name_key, data_list in sheets_data.items():
            df_report_sheet = pd.DataFrame(data_list)
            
            # Define default column order for each sheet type
            default_cols = ['RMA', 'S/N', 'SPC Code']
            if "Estimate Creation" in sheet_name_key: default_cols.extend(['Est. Complete Date'])
            elif "Reminder" in sheet_name_key and report_type != "EOD" : default_cols.extend(['Estimate Sent To Email', 'Estimate Sent Time']) # For Daily/Custom Reminder
            elif "Shipping" in sheet_name_key and report_type != "EOD" : pass # No extra for daily shipping
            elif "EOD" in sheet_name_key and "AdHoc" not in sheet_name_key: default_cols.extend(['Status', 'Original Task']) # For EOD main tasks
            elif "AdHoc" in sheet_name_key: default_cols = ['RMA', 'S/N', 'SPC Code', 'Shipped Time']


            if not df_report_sheet.empty:
                # Ensure all default_cols are present, add if missing
                for col in default_cols:
                    if col not in df_report_sheet.columns: df_report_sheet[col] = 'N/A'
                
                # Add any other columns from data_list that weren't in default_cols
                other_cols_present = [col for col in df_report_sheet.columns if col not in default_cols]
                final_cols_order = default_cols + other_cols_present
                # Ensure only existing columns are selected to prevent KeyErrors if a column was expected but not generated
                final_cols_order = [col for col in final_cols_order if col in df_report_sheet.columns]
                df_report_sheet = df_report_sheet[final_cols_order] # Reorder/select

                df_for_excel = df_report_sheet.astype(str) # Convert all to string for Excel

                df_for_excel.to_excel(writer, sheet_name=sheet_name_key, startrow=2, index=False, header=False)
                worksheet = writer.sheets[sheet_name_key]
                worksheet.merge_range(0, 0, 0, len(df_for_excel.columns)-1 if len(df_for_excel.columns)>0 else 0, f"{sheet_name_key} - Report Date: {report_data['date']}", title_format)
                worksheet.set_row(0, 30) 
                for col_num, value in enumerate(df_for_excel.columns.values): worksheet.write(2, col_num, value, header_format)
                for row_num in range(3, len(df_for_excel) + 3): 
                    for col_num in range(len(df_for_excel.columns)):
                        worksheet.write(row_num, col_num, df_for_excel.iloc[row_num-3, col_num], cell_format) # Value is already string
                for i, col_name_iter in enumerate(df_for_excel.columns): 
                    header_len = len(str(col_name_iter))
                    data_max_len = 0
                    if not df_for_excel[col_name_iter].empty:
                        try:
                            data_max_len = df_for_excel[col_name_iter].map(len).max()
                        except (ValueError, TypeError): data_max_len = 0 
                    column_width = max(data_max_len, header_len) + 2
                    worksheet.set_column(i, i, column_width)
            else: 
                worksheet = writer.book.add_worksheet(sheet_name_key) 
                worksheet.merge_range(0, 0, 0, 2, f"{sheet_name_key} - Report Date: {report_data['date']}", title_format)
                worksheet.write(2,0, "No items for this category.", cell_format)
    return output.getvalue()

def display_formatted_report(report_data, source="Newly Generated", report_key_suffix=""):
    st.markdown(f"### {source} Daily Status Report for: {report_data['date']}")
    st.markdown(f"**📋 Needs Estimate Creation (from items completed on {(datetime.strptime(report_data['date'], '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')}):**")
    if report_data['needs_estimate_creation']:
        for item in report_data['needs_estimate_creation']:
            st.markdown(f"- RMA: {item.get('RMA', 'N/A')}, S/N: {item.get('S/N', 'N/A')}, SPC: {item.get('SPC Code', 'N/A')} (Est. Complete: {item.get('Est. Complete Date', 'N/A')})")
    else: st.info("None for this category.")
    
    st.markdown(f"**📞 Needs Reminder (Estimate Sent 2 days prior to {report_data['date']}):**") 
    if report_data.get('needs_reminder'): 
        for item in report_data['needs_reminder']:
            st.markdown(f"- RMA: {item.get('RMA', 'N/A')}, S/N: {item.get('S/N', 'N/A')}, SPC: {item.get('SPC Code', 'N/A')}, Email: {item.get('Estimate Sent To Email', 'N/A')}, Sent Time: {item.get('Estimate Sent Time', 'N/A')}")
    else: st.info("None for this category.")

    st.markdown(f"**🚢 Needs Shipping (QA'd on {report_data['date']}):**")
    if report_data['needs_shipping']:
        for item in report_data['needs_shipping']:
            st.markdown(f"- RMA: {item.get('RMA', 'N/A')}, S/N: {item.get('S/N', 'N/A')}, SPC: {item.get('SPC Code', 'N/A')}")
    else: st.info("None for this category.")
    excel_bytes = create_excel_report_bytes(report_data, report_type=source) # Pass report type
    st.download_button(
        label=f"Download Report for {report_data['date']} (Excel)", data=excel_bytes,
        file_name=f"Daily_Status_Report_{report_data['date']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_report_{report_data['date']}_{report_key_suffix}" )
    st.markdown("---")

# --- Main Application ---
st.title("🛠️ Service Process Dashboard") 
st.markdown("Monitor and update service item statuses, including shipping.")

if 'first_load_complete' not in st.session_state: st.session_state.first_load_complete = False
if 'refresh_counter' not in st.session_state: st.session_state.refresh_counter = 0
if 'data_df' not in st.session_state:
    st.session_state.data_df = load_data_from_google_sheet()
    st.session_state.first_load_complete = True 
if 'newly_generated_reports_to_display' not in st.session_state: st.session_state.newly_generated_reports_to_display = []
if 'selected_archived_report_to_display' not in st.session_state: st.session_state.selected_archived_report_to_display = None
if 'custom_report_to_display' not in st.session_state: st.session_state.custom_report_to_display = None 
# if 'todays_tasks_for_eod_report' not in st.session_state: st.session_state.todays_tasks_for_eod_report = None # No longer primary source for EOD
# if 'todays_tasks_date' not in st.session_state: st.session_state.todays_tasks_date = None 
if 'end_of_day_summary_report' not in st.session_state: st.session_state.end_of_day_summary_report = None 
if 'selected_eod_summary_to_display' not in st.session_state: st.session_state.selected_eod_summary_to_display = None


if st.button("🔄 Refresh Data from Google Sheet"):
    load_data_from_google_sheet.clear(); st.session_state.data_df = load_data_from_google_sheet()
    get_archived_reports_from_gsheet.clear(archive_sheet_name=ARCHIVE_SHEET_NAME, expected_headers=ARCHIVE_SHEET_HEADERS) 
    get_archived_reports_from_gsheet.clear(archive_sheet_name=EOD_SUMMARY_ARCHIVE_SHEET_NAME, expected_headers=EOD_ARCHIVE_SHEET_HEADERS) 
    st.session_state.first_load_complete = False
    st.session_state.refresh_counter += 1 
    st.session_state.newly_generated_reports_to_display = [] 
    st.session_state.selected_archived_report_to_display = None 
    st.session_state.custom_report_to_display = None 
    st.session_state.end_of_day_summary_report = None 
    st.session_state.selected_eod_summary_to_display = None 
    st.rerun() 

data_df = st.session_state.data_df
archived_daily_reports_gsheet = get_archived_reports_from_gsheet(ARCHIVE_SHEET_NAME, ARCHIVE_SHEET_HEADERS) 
archived_eod_summaries_gsheet = get_archived_reports_from_gsheet(EOD_SUMMARY_ARCHIVE_SHEET_NAME, EOD_ARCHIVE_SHEET_HEADERS)


st.sidebar.markdown("---")
st.sidebar.header("📅 Daily Status Reports")
if st.sidebar.button("Generate Daily Status Report(s)", key=f"gen_daily_report_btn_{st.session_state.refresh_counter}"):
    st.session_state.newly_generated_reports_to_display = [] 
    st.session_state.selected_archived_report_to_display = None 
    st.session_state.custom_report_to_display = None 
    st.session_state.end_of_day_summary_report = None 
    st.session_state.selected_eod_summary_to_display = None

    if data_df.empty:
        st.sidebar.warning("No data loaded to generate reports.")
    else:
        last_gen_date_from_archive = get_last_report_date_from_archive(archived_daily_reports_gsheet)
        today = date.today()
        current_date_to_report = last_gen_date_from_archive + timedelta(days=1)
        reports_generated_this_run = []
        if current_date_to_report > today:
             st.sidebar.info("Daily reports are up to date according to archive.")
        else:
            while current_date_to_report <= today:
                report_data = generate_single_day_report_content(data_df, current_date_to_report)
                if save_report_to_gsheet_archive(report_data, ARCHIVE_SHEET_NAME, ARCHIVE_SHEET_HEADERS): 
                    reports_generated_this_run.append(report_data)
                    get_archived_reports_from_gsheet.clear(archive_sheet_name=ARCHIVE_SHEET_NAME, expected_headers=ARCHIVE_SHEET_HEADERS) 
                if current_date_to_report == today: break 
                current_date_to_report += timedelta(days=1)
                if (current_date_to_report - (last_gen_date_from_archive + timedelta(days=1))).days > 30 : 
                    st.sidebar.error("More than 30 days of reports to generate."); break
            if reports_generated_this_run:
                st.session_state.newly_generated_reports_to_display = reports_generated_this_run
                st.sidebar.success(f"{len(reports_generated_this_run)} daily report(s) generated and saved to Google Sheet archive.")
            elif last_gen_date_from_archive >= today : 
                st.sidebar.info("Daily report for today already in archive or no new days to report.")
            st.rerun() 

st.sidebar.markdown("---")
st.sidebar.header("🏁 End of Day Summary")
if st.sidebar.button("Generate End of Day Summary", key=f"gen_eod_summary_btn_{st.session_state.refresh_counter}"):
    st.session_state.newly_generated_reports_to_display = [] 
    st.session_state.selected_archived_report_to_display = None 
    st.session_state.custom_report_to_display = None 
    
    load_data_from_google_sheet.clear() 
    current_data_for_eod = load_data_from_google_sheet()

    todays_archived_report_for_eod = None
    today_str = date.today().strftime("%Y-%m-%d")
    # Make sure to use the correct archive list for daily reports
    for rpt in archived_daily_reports_gsheet: 
        if rpt['date'] == today_str:
            todays_archived_report_for_eod = rpt
            break
    
    if todays_archived_report_for_eod is None:
        st.sidebar.warning("Today's Daily Status Report has not been generated and archived yet. Please generate it first to establish the End of Day baseline.")
        st.session_state.end_of_day_summary_report = None
    elif current_data_for_eod.empty: 
        st.sidebar.warning("No data loaded to generate end-of-day summary.")
        st.session_state.end_of_day_summary_report = None
    else:
        summary_data = {"date": today_str, "estimate_tasks": [], "shipping_tasks": [], "reminder_tasks": [], "adhoc_shipped_today": []} 
        
        tasks_from_daily_report = todays_archived_report_for_eod 
        
        for task_type, task_list_key, status_col, completion_value, task_desc_template_base in [
            ("estimate_tasks", "needs_estimate_creation", "Estimate Sent To Email", "n/a", "Create/Send Estimate (Est. Complete: {Est. Complete Date})"),
            ("reminder_tasks", "needs_reminder", "Reminder Completed", "yes", "Send Reminder (Est. Sent: {Estimate Sent Time})"),
            ("shipping_tasks", "needs_shipping", "Shipped", "yes", "Ship Item (QA'd on " + tasks_from_daily_report.get('date', 'N/A') + ")")
        ]:
            for task in tasks_from_daily_report.get(task_list_key, []): 
                rma, sn = task.get("RMA"), task.get("S/N")
                spc_code = task.get("SPC Code", "N/A") 
                rma_str = str(rma).strip().lower(); sn_str = str(sn).strip().lower()
                record_df = current_data_for_eod[
                    (current_data_for_eod['RMA'].astype(str).str.strip().str.lower() == rma_str) &
                    (current_data_for_eod['S/N'].astype(str).str.strip().str.lower() == sn_str) ]
                status = "Pending" 
                if not record_df.empty:
                    current_status_val = record_df.iloc[0][status_col].lower()
                    if (status_col == "Estimate Sent To Email" and current_status_val != 'n/a') or \
                       (status_col != "Estimate Sent To Email" and current_status_val == completion_value):
                        status = "Completed"
                
                task_description = f"Task for RMA {rma}, S/N {sn}" # Default
                if task_type == "estimate_tasks": task_description = f"Create/Send Estimate (Est. Complete: {task.get('Est. Complete Date', 'N/A')})"
                elif task_type == "reminder_tasks": task_description = f"Send Reminder (Est. Sent: {task.get('Estimate Sent Time', 'N/A')})"
                elif task_type == "shipping_tasks": task_description = f"Ship Item (QA'd on {tasks_from_daily_report.get('date', 'N/A')})"
                
                summary_data[task_type].append({"RMA": rma, "S/N": sn, "SPC Code": spc_code, "Status": status, "Original Task": task_description})

        # Identify Ad-hoc Shipped Today
        adhoc_shipped_df = current_data_for_eod[
            (current_data_for_eod['Shipped'].astype(str).str.lower() == 'yes') &
            (pd.to_datetime(current_data_for_eod['Shipped Time'], errors='coerce').dt.date == date.today())
        ]
        daily_shipping_rmas_sns = [(str(t.get("RMA")).strip().lower(), str(t.get("S/N")).strip().lower()) for t in tasks_from_daily_report.get("needs_shipping", [])]
        for _, row in adhoc_shipped_df.iterrows():
            rma_str = str(row['RMA']).strip().lower(); sn_str = str(row['S/N']).strip().lower()
            if (rma_str, sn_str) not in daily_shipping_rmas_sns: 
                summary_data["adhoc_shipped_today"].append({
                    "RMA": row['RMA'], "S/N": row['S/N'], "SPC Code": row.get('SPC Code', 'N/A'), 
                    "Shipped Time": pd.to_datetime(row['Shipped Time']).strftime('%Y-%m-%d %H:%M') if pd.notna(row['Shipped Time']) else 'N/A' })

        st.session_state.end_of_day_summary_report = summary_data
        if save_report_to_gsheet_archive(summary_data, EOD_SUMMARY_ARCHIVE_SHEET_NAME, EOD_ARCHIVE_SHEET_HEADERS):
            st.sidebar.success("End of Day Summary Generated and Archived to Google Sheet.")
        else: st.sidebar.warning("End of Day Summary Generated (already exists in archive or error saving).")
        st.rerun()

# --- Display Sections ---
if st.session_state.newly_generated_reports_to_display:
    st.markdown("---"); st.subheader("✨ Newly Generated Daily Status Report(s)")
    for i, report in enumerate(st.session_state.newly_generated_reports_to_display):
        display_formatted_report(report, source="Newly Generated", report_key_suffix=f"new_{i}")
    if st.button("Clear Newly Generated Reports View", key="clear_new_reports"):
        st.session_state.newly_generated_reports_to_display = []; st.rerun()
    st.markdown("---")

if st.session_state.selected_archived_report_to_display: # For Daily Status Archive
    st.markdown("---")
    display_formatted_report(st.session_state.selected_archived_report_to_display, source="Archived Daily Report", report_key_suffix="archive_daily_disp")
    if st.button("Close Archived Daily Report View", key="close_archive_daily_view"):
        st.session_state.selected_archived_report_to_display = None; st.rerun()
    st.markdown("---")

if st.session_state.custom_report_to_display:
    st.markdown("---")
    display_formatted_report(st.session_state.custom_report_to_display, source="Custom Date Report", report_key_suffix="custom_disp")
    if st.button("Clear Custom Report View", key="clear_custom_report"):
        st.session_state.custom_report_to_display = None; st.rerun()
    st.markdown("---")

if st.session_state.get('end_of_day_summary_report'): # For Live EOD Summary
    eod_summary = st.session_state.end_of_day_summary_report
    st.markdown("---"); st.subheader(f"🏁 End of Day Summary for: {eod_summary['date']}")
    eod_display_cols = ["RMA", "S/N", "SPC Code", "Status", "Original Task"]; adhoc_shipped_cols = ["RMA", "S/N", "SPC Code", "Shipped Time"]
    st.markdown("**Estimate Creation Task Summary:**")
    if eod_summary['estimate_tasks']: 
        eod_est_df = pd.DataFrame(eod_summary['estimate_tasks'])
        for col in eod_display_cols:
            if col not in eod_est_df.columns: eod_est_df[col] = "N/A"
        st.dataframe(eod_est_df[eod_display_cols], use_container_width=True)
    else: st.info("No estimate creation tasks were on today's daily report or daily report not generated for today.")
    st.markdown("**Reminder Task Summary:**") 
    if eod_summary.get('reminder_tasks'): 
        eod_rem_df = pd.DataFrame(eod_summary['reminder_tasks'])
        for col in eod_display_cols: 
            if col not in eod_rem_df.columns: eod_rem_df[col] = "N/A"
        st.dataframe(eod_rem_df[eod_display_cols], use_container_width=True)
    else: st.info("No reminder tasks were on today's daily report or daily report not generated for today.")
    st.markdown("**Shipping Task Summary (from Daily Report):**")
    if eod_summary['shipping_tasks']: 
        eod_ship_df = pd.DataFrame(eod_summary['shipping_tasks'])
        for col in eod_display_cols: 
            if col not in eod_ship_df.columns: eod_ship_df[col] = "N/A"
        st.dataframe(eod_ship_df[eod_display_cols], use_container_width=True)
    else: st.info("No shipping tasks were on today's daily report or daily report not generated for today.")
    st.markdown("**Ad-hoc Shipped Today (not on initial daily report):**") 
    if eod_summary.get('adhoc_shipped_today'):
        eod_adhoc_df = pd.DataFrame(eod_summary['adhoc_shipped_today'])
        for col in adhoc_shipped_cols: 
            if col not in eod_adhoc_df.columns: eod_adhoc_df[col] = "N/A"
        st.dataframe(eod_adhoc_df[adhoc_shipped_cols], use_container_width=True)
    else: st.info("No additional items were marked as shipped today outside of the daily report tasks.")
    eod_output = BytesIO()
    with pd.ExcelWriter(eod_output, engine='xlsxwriter') as writer:
        workbook = writer.book
        header_format = workbook.add_format({'bold': True, 'text_wrap': False, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
        cell_format = workbook.add_format({'border': 1})
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
        eod_sheets_data = { 
            "EOD Estimate Tasks": eod_summary.get("estimate_tasks", []),
            "EOD Reminder Tasks": eod_summary.get("reminder_tasks", []), 
            "EOD Shipping Tasks": eod_summary.get("shipping_tasks", []),
            "EOD AdHoc Shipped": eod_summary.get("adhoc_shipped_today", []) 
        }
        for sheet_name_key, data_list in eod_sheets_data.items():
            df_eod_sheet = pd.DataFrame(data_list)
            current_display_cols = eod_display_cols if "AdHoc" not in sheet_name_key else adhoc_shipped_cols
            if not df_eod_sheet.empty:
                for col in current_display_cols: 
                    if col not in df_eod_sheet.columns: df_eod_sheet[col] = "N/A"
                df_eod_sheet = df_eod_sheet[current_display_cols] 
            if not df_eod_sheet.empty:
                df_eod_sheet.to_excel(writer, sheet_name=sheet_name_key, startrow=2, index=False, header=False)
                ws = writer.sheets[sheet_name_key]
                ws.merge_range(0,0,0, len(df_eod_sheet.columns)-1 if len(df_eod_sheet.columns)>0 else 0, f"{sheet_name_key} - {eod_summary['date']}", title_format)
                ws.set_row(0,30)
                for cn, val in enumerate(df_eod_sheet.columns.values): ws.write(2, cn, val, header_format)
                for rn in range(3, len(df_eod_sheet)+3):
                    for cn_idx in range(len(df_eod_sheet.columns)): ws.write(rn, cn_idx, df_eod_sheet.iloc[rn-3, cn_idx], cell_format)
                for i, col in enumerate(df_eod_sheet.columns):
                    col_len = max(df_eod_sheet[col].astype(str).map(len).max(), len(col)) + 2 if not df_eod_sheet[col].empty else len(col)+2
                    ws.set_column(i,i,col_len)
            else:
                ws = writer.book.add_worksheet(sheet_name_key)
                ws.merge_range(0,0,0,2, f"{sheet_name_key} - {eod_summary['date']}", title_format)
                ws.write(2,0, f"No tasks for this category on {eod_summary['date']}.", cell_format)
    st.download_button(label=f"Download End of Day Summary ({eod_summary['date']})", data=eod_output.getvalue(),
                       file_name=f"EndOfDay_Summary_{eod_summary['date']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=f"download_eod_summary_{eod_summary['date']}_{st.session_state.refresh_counter}")
    if st.button("Clear End of Day Summary View", key=f"clear_eod_summary_{st.session_state.refresh_counter}"):
        st.session_state.end_of_day_summary_report = None; st.rerun()
    st.markdown("---")

# --- Display Selected Archived EOD Summary ---
if st.session_state.get('selected_eod_summary_to_display'):
    eod_summary_to_show = st.session_state.selected_eod_summary_to_display
    st.markdown("---") 
    st.subheader(f"Archived End of Day Summary for: {eod_summary_to_show.get('date', 'N/A')}")
    
    eod_display_cols = ["RMA", "S/N", "SPC Code", "Status", "Original Task"]
    adhoc_shipped_cols = ["RMA", "S/N", "SPC Code", "Shipped Time"]

    st.markdown("**Estimate Creation Task Summary:**")
    if eod_summary_to_show.get('estimate_tasks'): 
        eod_est_df = pd.DataFrame(eod_summary_to_show['estimate_tasks'])
        for col in eod_display_cols:
            if col not in eod_est_df.columns: eod_est_df[col] = "N/A"
        st.dataframe(eod_est_df[eod_display_cols], use_container_width=True)
    else: st.info("No estimate creation tasks in this archived summary.")
    
    st.markdown("**Reminder Task Summary:**") 
    if eod_summary_to_show.get('reminder_tasks'): 
        eod_rem_df = pd.DataFrame(eod_summary_to_show['reminder_tasks'])
        for col in eod_display_cols: 
            if col not in eod_rem_df.columns: eod_rem_df[col] = "N/A"
        st.dataframe(eod_rem_df[eod_display_cols], use_container_width=True)
    else: st.info("No reminder tasks in this archived summary.")
    
    st.markdown("**Shipping Task Summary (from Daily Report):**")
    if eod_summary_to_show.get('shipping_tasks'): 
        eod_ship_df = pd.DataFrame(eod_summary_to_show['shipping_tasks'])
        for col in eod_display_cols: 
            if col not in eod_ship_df.columns: eod_ship_df[col] = "N/A"
        st.dataframe(eod_ship_df[eod_display_cols], use_container_width=True)
    else: st.info("No shipping tasks from daily report in this archived summary.")

    st.markdown("**Ad-hoc Shipped Today (not on initial daily report):**")
    if eod_summary_to_show.get('adhoc_shipped_today'):
        eod_adhoc_df = pd.DataFrame(eod_summary_to_show['adhoc_shipped_today'])
        for col in adhoc_shipped_cols: 
            if col not in eod_adhoc_df.columns: eod_adhoc_df[col] = "N/A"
        st.dataframe(eod_adhoc_df[adhoc_shipped_cols], use_container_width=True)
    else: st.info("No additional items were marked as shipped ad-hoc in this archived summary.")

    excel_bytes_eod_archive = create_excel_report_bytes(eod_summary_to_show, report_type="EOD")
    st.download_button(
        label=f"Download EOD Summary for {eod_summary_to_show.get('date', 'N/A')} (Excel)", 
        data=excel_bytes_eod_archive,
        file_name=f"EOD_Summary_{eod_summary_to_show.get('date', 'N/A')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_archived_eod_{eod_summary_to_show.get('date', 'N/A')}_{st.session_state.refresh_counter}"
    )
    if st.button("Close Archived EOD Summary View", key=f"close_archived_eod_view_{st.session_state.refresh_counter}"):
        st.session_state.selected_eod_summary_to_display = None
        st.rerun()
    st.markdown("---")


if not data_df.empty:
    st.subheader("📊 Key Metrics")
    display_kpis(data_df.copy()); st.markdown("---")
    st.subheader("⚠️ Overdue Estimates Report (Pending Approval > 3 Days)")
    overdue_estimates_df = identify_overdue_estimates(data_df, days_threshold=3) 
    if not overdue_estimates_df.empty:
        st.warning("The following estimates were completed more than 3 days ago and are still pending approval:")
        overdue_estimates_display_cols = ['RMA', 'S/N', 'SPC Code', 'Estimate Complete Time', 'Days Pending Approval', 
                                          'Estimate Sent To Email', 'Estimate Sent Time', 
                                          'Reminder Completed', 'Reminder Completed Time', 'Reminder Contact Method', BC_LINK_COL_NAME]
        for col in overdue_estimates_display_cols:
            if col not in overdue_estimates_df.columns: overdue_estimates_df[col] = None 
        st.dataframe(overdue_estimates_df[overdue_estimates_display_cols], use_container_width=True,
            column_config={ BC_LINK_COL_NAME: st.column_config.LinkColumn(label="Business Central", display_text="Open RMA")},
            column_order=overdue_estimates_display_cols )
    else: st.success("✅ No estimates are currently overdue for approval beyond 3 days.")
    st.markdown("---")
    
    st.subheader("🗣️ Overdue Reminders Report (Estimate Sent > 2 Days, Reminder Not Done)") 
    overdue_reminders_df = identify_overdue_reminders(data_df, days_threshold=2)
    if not overdue_reminders_df.empty:
        st.info("The following items had estimates sent more than 2 days ago and are pending a reminder:")
        overdue_reminders_display_cols = ['RMA', 'S/N', 'SPC Code', 'Estimate Sent To Email', 'Estimate Sent Time', 'Days Pending Reminder', 'Reminder Contact Method', BC_LINK_COL_NAME] 
        for col in overdue_reminders_display_cols:
            if col not in overdue_reminders_df.columns:
                overdue_reminders_df[col] = None
        st.dataframe(
            overdue_reminders_df[overdue_reminders_display_cols],
            use_container_width=True,
            column_config={
                BC_LINK_COL_NAME: st.column_config.LinkColumn(label="Business Central", display_text="Open RMA")
            },
            column_order=overdue_reminders_display_cols
        )
    else:
        st.success("✅ No items are currently overdue for a reminder beyond 2 days.")
    st.markdown("---")


    st.subheader("🚚 Overdue for Shipping Report (QA Approved > 1 Day, Not Shipped)")
    overdue_shipping_df = identify_overdue_for_shipping(data_df, days_threshold=1)
    if not overdue_shipping_df.empty:
        st.error("The following items are QA Approved for more than 1 day and are pending shipment:") 
        overdue_shipping_display_cols = ['RMA', 'S/N', 'SPC Code', 'QA Approved Time', 'Days Pending Shipping', BC_LINK_COL_NAME] 
        if BC_LINK_COL_NAME not in overdue_shipping_df.columns: overdue_shipping_df[BC_LINK_COL_NAME] = None
        st.dataframe(overdue_shipping_df[overdue_shipping_display_cols], use_container_width=True,
            column_config={BC_LINK_COL_NAME: st.column_config.LinkColumn(label="Business Central", display_text="Open RMA")},
            column_order=overdue_shipping_display_cols)
    else: st.success("✅ No items are currently overdue for shipping beyond 1 day.")
    st.markdown("---")

    st.subheader("🗂️ Daily Status Report Archive") 
    if archived_daily_reports_gsheet: 
        report_dates = sorted(list(set(r['date'] for r in archived_daily_reports_gsheet)), reverse=True)
        available_months = sorted(list(set(datetime.strptime(d, "%Y-%m-%d").strftime("%Y-%m") for d in report_dates)), reverse=True)
        if available_months:
            selected_month_archive = st.selectbox("View Daily Reports for Month:", ["All"] + available_months, key="archive_daily_month_select")
            reports_to_list = [r for r in archived_daily_reports_gsheet if selected_month_archive == "All" or datetime.strptime(r['date'], "%Y-%m-%d").strftime("%Y-%m") == selected_month_archive]
            if reports_to_list:
                for i, report_data_item in enumerate(reports_to_list): 
                    col1, col2 = st.columns([3,1])
                    with col1: st.markdown(f"**Daily Report for: {report_data_item['date']}**")
                    with col2:
                        if st.button("View/Download Daily Report", key=f"view_archive_daily_{report_data_item['date']}_{i}"): 
                            st.session_state.selected_archived_report_to_display = report_data_item
                            st.session_state.newly_generated_reports_to_display = []; st.session_state.custom_report_to_display = None; st.session_state.end_of_day_summary_report = None; st.session_state.selected_eod_summary_to_display = None; st.rerun()
            else: st.info(f"No daily reports found for {selected_month_archive} in the Google Sheet archive.")
        else: st.info("No archived daily reports available in the Google Sheet yet.")
    else: st.info("No archived daily reports available in the Google Sheet yet. Generate reports using the button in the sidebar.")
    st.markdown("---")

    st.subheader("🗂️ End of Day Summary Archive")
    if archived_eod_summaries_gsheet:
        eod_report_dates = sorted(list(set(r['date'] for r in archived_eod_summaries_gsheet)), reverse=True)
        eod_available_months = sorted(list(set(datetime.strptime(d, "%Y-%m-%d").strftime("%Y-%m") for d in eod_report_dates)), reverse=True)
        if eod_available_months:
            selected_eod_month_archive = st.selectbox("View EOD Summaries for Month:", ["All"] + eod_available_months, key="archive_eod_month_select")
            eod_summaries_to_list = [r for r in archived_eod_summaries_gsheet if selected_eod_month_archive == "All" or datetime.strptime(r['date'], "%Y-%m-%d").strftime("%Y-%m") == selected_eod_month_archive]
            if eod_summaries_to_list:
                for i, eod_summary_item in enumerate(eod_summaries_to_list):
                    col1, col2 = st.columns([3,1])
                    with col1: st.markdown(f"**EOD Summary for: {eod_summary_item['date']}**")
                    with col2:
                        if st.button("View/Download EOD Summary", key=f"view_archive_eod_{eod_summary_item['date']}_{i}"):
                            st.session_state.selected_eod_summary_to_display = eod_summary_item 
                            st.session_state.newly_generated_reports_to_display = []
                            st.session_state.selected_archived_report_to_display = None
                            st.session_state.custom_report_to_display = None
                            st.session_state.end_of_day_summary_report = None 
                            st.rerun()
            else: st.info(f"No EOD summaries found for {selected_eod_month_archive} in the Google Sheet archive.")
        else: st.info("No archived EOD summaries available in the Google Sheet yet.")
    else: st.info("No archived EOD summaries available in the Google Sheet yet.")
    st.markdown("---")


    st.sidebar.header("🔍 Filter Options")
    filtered_df = data_df.copy() 
    for col_name, search_label in [('RMA', "RMA"), ('S/N', "S/N"), ('Part Number', "Part Number"), ('SPC Code', "SPC Code")]:
        if col_name in filtered_df.columns:
            search_term = st.sidebar.text_input(f"Search by {search_label}", key=f"search_{col_name}_{st.session_state.refresh_counter}") 
            if search_term: filtered_df = filtered_df[filtered_df[col_name].astype(str).str.contains(search_term, case=False, na=False)]
    status_columns_to_filter = {
        'Estimate Complete': 'Estimate Complete', 'Estimate Approved': 'Estimate Approved',
        'Reminder Completed': 'Reminder Completed', 'QA Approved': 'QA Approved', 'Shipped': 'Shipped' }
    for display_name, col_name in status_columns_to_filter.items():
        if col_name in data_df.columns: 
            unique_values = ['All'] + sorted(list(set(val for val in data_df[col_name].astype(str).unique() if val and val.strip() != '' and val != 'N/A'))) 
            if 'N/A' in data_df[col_name].astype(str).unique(): unique_values.insert(1, "N/A") 
            if 'No' in data_df[col_name].astype(str).unique() and 'No' not in unique_values : unique_values.insert(1, "No") 
            default_index = 0 
            if not st.session_state.first_load_complete:
                if col_name == "Shipped" and "N/A" in unique_values: default_index = unique_values.index("N/A")
                elif col_name == "Reminder Completed" and "No" in unique_values: default_index = unique_values.index("No")
            current_key = f"select_{col_name}_{st.session_state.refresh_counter}"
            selected_status = st.sidebar.selectbox(f"Filter by {display_name}", unique_values, key=current_key, index=default_index)
            if selected_status != "All": 
                if col_name in filtered_df.columns: filtered_df = filtered_df[filtered_df[col_name].astype(str) == selected_status]
    st.sidebar.markdown("---"); st.sidebar.subheader("Date Range Filters")
    date_filter_columns_to_filter = {
        'Estimate Complete Time': 'Estimate Complete Time', 'Estimate Approved Time': 'Estimate Approved Time',
        'Estimate Sent Time': 'Estimate Sent Time', 'Reminder Completed Time': 'Reminder Completed Time', 
        'QA Approved Time': 'QA Approved Time', 'Shipped Time': 'Shipped Time' }
    for display_name, col_name in date_filter_columns_to_filter.items():
        min_val_for_widget_setup = None; max_val_for_widget_setup = None; can_setup_widget = False
        if col_name in data_df.columns and pd.api.types.is_datetime64_any_dtype(data_df[col_name]):
            original_col_for_widget_params = data_df[col_name].dropna() 
            if not original_col_for_widget_params.empty:
                min_val_for_widget_setup = original_col_for_widget_params.min().date(); max_val_for_widget_setup = original_col_for_widget_params.max().date()
                can_setup_widget = True
        if can_setup_widget:
            current_key_date = f"date_range_{col_name}_{st.session_state.refresh_counter}"
            current_date_range_selection = st.sidebar.date_input(f"Filter by {display_name}", value=[], 
                min_value=min_val_for_widget_setup, max_value=max_val_for_widget_setup, key=current_key_date) 
            if current_date_range_selection and len(current_date_range_selection) == 2: 
                if col_name in filtered_df.columns and pd.api.types.is_datetime64_any_dtype(filtered_df[col_name]):
                    start_date_selected, end_date_selected = current_date_range_selection
                    start_datetime_selected = pd.to_datetime(start_date_selected); end_datetime_selected = pd.to_datetime(end_date_selected).replace(hour=23, minute=59, second=59) 
                    condition = ((filtered_df[col_name] >= start_datetime_selected) & (filtered_df[col_name] <= end_datetime_selected) & (filtered_df[col_name].notna()) )
                    filtered_df = filtered_df[condition]
    if not st.session_state.first_load_complete: st.session_state.first_load_complete = True

    st.sidebar.markdown("---"); st.sidebar.header("📝 Log Estimate Sent")
    if all(c in data_df.columns for c in ['RMA', 'S/N', 'Estimate Complete', 'Estimate Sent To Email']): 
        eligible_estimate_sent_df = data_df[ 
            (data_df['Estimate Complete'].astype(str).str.lower() == 'yes') & 
            (data_df['Estimate Sent To Email'].astype(str).str.lower() == 'n/a') 
        ]
        if not eligible_estimate_sent_df.empty:
            options = ["Select item..."] + [f"{rma} - S/N: {sn}" for rma, sn in zip(eligible_estimate_sent_df['RMA'], eligible_estimate_sent_df['S/N'])]
            selected_item_est_sent = st.sidebar.selectbox("Item (Est. Complete, Not Sent)", options, index=0, key=f"est_sent_item_selector_{st.session_state.refresh_counter}")
            if selected_item_est_sent and selected_item_est_sent != "Select item...":
                rma_est_sent, sn_part_est_sent = selected_item_est_sent.split(" - S/N: "); sn_est_sent = sn_part_est_sent.strip()
                sent_to_email = st.sidebar.text_input("Sent To Email Address", key=f"est_sent_email_input_{st.session_state.refresh_counter}")
                sent_date_val = st.sidebar.date_input("Estimate Sent Date", value=date.today(), key=f"est_sent_date_input_{st.session_state.refresh_counter}")
                if st.sidebar.button("Mark Estimate Sent", key=f"mark_est_sent_button_{st.session_state.refresh_counter}"):
                    if rma_est_sent and sn_est_sent and sent_to_email and sent_date_val:
                        if "@" not in sent_to_email or "." not in sent_to_email: st.sidebar.error("Please enter a valid email address.")
                        else:
                            success = update_estimate_sent_details_in_gsheet(rma_est_sent, sn_est_sent, sent_to_email, sent_date_val)
                            if success: load_data_from_google_sheet.clear(); st.session_state.data_df = load_data_from_google_sheet(); st.session_state.first_load_complete = False; st.session_state.refresh_counter +=1; st.sidebar.success("Estimate sent details updated!"); st.rerun()
                    else: st.sidebar.warning("Please select item, enter email, and date.")
        else: st.sidebar.info("No items currently eligible for marking as estimate sent.")
    
    st.sidebar.markdown("---"); st.sidebar.header("📞 Log Reminder")
    if all(c in data_df.columns for c in ['RMA', 'S/N', 'Estimate Sent To Email', 'Reminder Completed']):
        eligible_reminder_df = data_df[ (data_df['Estimate Sent To Email'].astype(str).str.lower() != 'n/a') & (data_df['Reminder Completed'].astype(str).str.lower().isin(['no', 'n/a'])) ]
        if not eligible_reminder_df.empty:
            options = ["Select item..."] + [f"{rma} - S/N: {sn}" for rma, sn in zip(eligible_reminder_df['RMA'], eligible_reminder_df['S/N'])]
            selected_item_reminder = st.sidebar.selectbox("Item (Estimate Sent, Reminder Pending)", options, index=0, key=f"reminder_item_selector_{st.session_state.refresh_counter}")
            if selected_item_reminder and selected_item_reminder != "Select item...":
                rma_reminder, sn_part_reminder = selected_item_reminder.split(" - S/N: "); sn_reminder = sn_part_reminder.strip()
                contact_method_options = ["Email", "Phone Call", "Text", "Other"]
                reminder_contact_method = st.sidebar.selectbox("Reminder Contact Method", contact_method_options, key=f"reminder_contact_method_{st.session_state.refresh_counter}")
                reminder_date_val = st.sidebar.date_input("Reminder Date", value=date.today(), key=f"reminder_date_input_{st.session_state.refresh_counter}")
                if st.sidebar.button("Mark Reminder Completed", key=f"mark_reminder_button_{st.session_state.refresh_counter}"):
                    if rma_reminder and sn_reminder and reminder_date_val and reminder_contact_method: 
                        success = update_reminder_details_in_gsheet(rma_reminder, sn_reminder, reminder_date_val, reminder_contact_method)
                        if success: load_data_from_google_sheet.clear(); st.session_state.data_df = load_data_from_google_sheet(); st.session_state.first_load_complete = False; st.session_state.refresh_counter +=1; st.sidebar.success("Reminder details updated!"); st.rerun()
                    else: st.sidebar.warning("Please select an item, contact method, and date.")
        else: st.sidebar.info("No items currently eligible for reminder logging.")

    st.sidebar.markdown("---"); st.sidebar.header("📦 Update Shipped Status") 
    if all(c in data_df.columns for c in ['RMA', 'S/N', 'Shipped']):
        unshipped_items_df = data_df[data_df['Shipped'].astype(str).str.lower().isin(['no', 'n/a'])]
        if not unshipped_items_df.empty:
            unshipped_options = ["Select an item..."] + [f"{rma} - S/N: {sn}" for rma, sn in zip(unshipped_items_df['RMA'], unshipped_items_df['S/N'])]
            selected_item_str = st.sidebar.selectbox("Select Item to Mark as Shipped (RMA - S/N)", options=unshipped_options, index=0, key=f"shipped_item_selector_{st.session_state.refresh_counter}")
            if selected_item_str and selected_item_str != "Select an item...":
                try:
                    rma_to_update, sn_part = selected_item_str.split(" - S/N: "); sn_to_update = sn_part.strip()
                    shipped_date_val = st.sidebar.date_input("Shipped Date", value=date.today(), key=f"shipped_date_input_{st.session_state.refresh_counter}") 
                    if st.sidebar.button("Mark as Shipped", key=f"mark_shipped_button_{st.session_state.refresh_counter}"):
                        if rma_to_update and sn_to_update and shipped_date_val:
                            success = update_shipped_status_in_gsheet(rma_to_update, sn_to_update, shipped_date_val)
                            if success: load_data_from_google_sheet.clear(); st.session_state.data_df = load_data_from_google_sheet(); st.session_state.first_load_complete = False; st.session_state.refresh_counter +=1; st.sidebar.success("Update successful! Data refreshed."); st.rerun() 
                            else: st.sidebar.error("Update failed. Check logs or details above.")
                        else: st.sidebar.warning("Please select an item and a valid shipped date.")
                except ValueError: st.sidebar.error("Invalid item format selected. Please re-select.")
        elif not data_df.empty : st.sidebar.info("All available items are marked as shipped.")


    st.subheader("Filtered Data View")
    st.markdown(f"Displaying **{len(filtered_df)}** records out of **{len(data_df) if not data_df.empty else 0}** total records.")
    if not filtered_df.empty:
        df_for_display = filtered_df.copy()
        df_for_display[BC_LINK_COL_NAME] = df_for_display.apply(
            lambda row: f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(str(row['RMA']))}%27"
            if pd.notna(row['RMA']) and str(row['RMA']).strip() != 'N/A' and str(row['RMA']).strip() != "" else None, axis=1 )
        display_cols_order = EXPECTED_COLUMN_ORDER[:] 
        if 'RMA' in display_cols_order: display_cols_order.insert(display_cols_order.index('RMA') + 1, BC_LINK_COL_NAME)
        else: display_cols_order.append(BC_LINK_COL_NAME)
        final_display_columns = [col for col in display_cols_order if col in df_for_display.columns]
        st.dataframe(df_for_display[final_display_columns], use_container_width=True,
            column_config={ BC_LINK_COL_NAME: st.column_config.LinkColumn(label="Business Central", display_text="Open RMA")})
    else: st.warning("No data matches the current filter criteria or no data loaded.")

    if not filtered_df.empty:
        st.sidebar.markdown("---"); st.sidebar.subheader("Download Data")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            display_cols_download = [col for col in EXPECTED_COLUMN_ORDER if col in filtered_df.columns]
            df_to_export = filtered_df[display_cols_download].copy()
            for col in df_to_export.columns:
                if pd.api.types.is_datetime64_any_dtype(df_to_export[col]):
                    if df_to_export[col].dt.tz is not None: df_to_export[col] = df_to_export[col].dt.tz_localize(None)
            df_to_export.to_excel(writer, index=False, sheet_name='ServiceData')
        excel_data = output.getvalue()
        st.sidebar.download_button(label="Download Filtered Data as XLSX", data=excel_data,
            file_name='filtered_service_data.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='download_xlsx')
else:
    st.info("No data to display. Please ensure the Google Sheet is accessible, contains data with headers for all expected columns, and 'Credentials.json' is correctly set up.")

st.markdown("---")
st.markdown("Built with ❤️ using [Streamlit](https://streamlit.io) and Google Sheets")
