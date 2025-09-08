import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta 
import gspread 
from oauth2client.service_account import ServiceAccountCredentials 
from io import BytesIO 
import urllib.parse 

# --- Page Configuration ---
st.set_page_config(
    page_title="Service Data Dashboard", 
    page_icon="🚚", 
    layout="wide",
)

# --- Constants for Google Sheets ---
GSHEET_NAME = "Estimate form"
WORKSHEET_INDEX = 1 
CREDS_FILE = "Credentials.json" # Make sure this file is in the same directory
EXPECTED_COLUMN_ORDER = [
    "RMA", "SPC Code", "Part Number", "S/N", "Description", 
    "Fault Comments", "Resolution Comments", "Sender", 
    "Estimate Complete Time", "Estimate Complete", 
    "Estimate Approved", "Estimate Approved Time",
    "Estimate Sent To Email", "Estimate Sent Time", 
    "Reminder Completed", "Reminder Completed Time", 
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
            st.warning(f"No data (not even headers) found in Google Sheet '{sheet_name}', worksheet index {worksheet_index}.")
            return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) 
            
        headers_from_sheet = all_values[0]
        data_rows = all_values[1:]
        
        temp_df = pd.DataFrame(data_rows, columns=headers_from_sheet)
        df = pd.DataFrame(columns=EXPECTED_COLUMN_ORDER)

        for col in EXPECTED_COLUMN_ORDER:
            if col in temp_df.columns:
                df[col] = temp_df[col] 
            else: 
                 st.warning(f"Expected column '{col}' not found in Google Sheet. Initializing as empty/default.")
                 if "Time" in col: df[col] = pd.NaT
                 elif col in ALL_STATUS_COLUMNS: df[col] = "No"
                 elif col == "Estimate Sent To Email": df[col] = "N/A" 
                 else: df[col] = "N/A" 
        
        df = df[EXPECTED_COLUMN_ORDER] 

        string_cols_to_process = ['RMA', 'S/N', 'Part Number', 'SPC Code', 
                                   'Description', 'Fault Comments', 'Resolution Comments', 
                                   'Sender', 'Estimate Sent To Email'] + ALL_STATUS_COLUMNS
        for col in string_cols_to_process:
            if col in df.columns:
                df[col] = df[col].astype(str) 
                if col in ALL_STATUS_COLUMNS:
                     df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'No') 
                elif col == "Estimate Sent To Email":
                     df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'N/A')
                else: 
                     df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'N/A')

        for col in ALL_TIME_COLUMNS:
            if col in df.columns:
                df[col] = df[col].replace('N/A', None) 
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        return df

    except FileNotFoundError:
        st.error(f"Error: Credentials file '{creds_file}' not found. Please ensure it's in the correct path.")
        return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) 
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"Error: Google Sheet '{sheet_name}' not found. Please check the name and permissions.")
        return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) 
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Error: Worksheet with index {worksheet_index} not found in '{sheet_name}'.")
        return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) 
    except Exception as e:
        st.error(f"An error occurred while loading data from Google Sheets: {type(e).__name__} - {e}")
    return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) 


def find_row_in_gsheet(worksheet, rma_to_find, sn_to_find, headers):
    try:
        rma_col_idx = headers.index("RMA") 
        sn_col_idx = headers.index("S/N") 
    except ValueError:
        st.error("RMA or S/N column not found in sheet headers during row search.")
        return -1
    all_data_values = worksheet.get_all_values() 
    for i, row_values in enumerate(all_data_values[1:], start=2): 
        rma_val = row_values[rma_col_idx] if len(row_values) > rma_col_idx else None
        sn_val = row_values[sn_col_idx] if len(row_values) > sn_col_idx else None
        if rma_val == rma_to_find and sn_val == sn_to_find:
            return i 
    return -1

def update_gsheet_cells(worksheet, updates_list):
    try:
        worksheet.batch_update(updates_list)
        return True
    except Exception as e:
        st.error(f"An error occurred during Google Sheet batch update: {e}")
        return False

def gsheet_update_wrapper(update_function, *args):
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open(GSHEET_NAME)
        worksheet = spreadsheet.get_worksheet(WORKSHEET_INDEX)
        headers = worksheet.row_values(1)
        if not headers:
            st.error("Could not read headers from Google Sheet. Update failed.")
            return False
        return update_function(worksheet, headers, *args)
    except Exception as e:
        st.error(f"General error during Google Sheet operation: {type(e).__name__} - {e}")
        return False

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

def _update_reminder_in_sheet(worksheet, headers, rma, sn, reminder_date_obj):
    reminder_status_col_name = "Reminder Completed"; reminder_time_col_name = "Reminder Completed Time"
    try:
        reminder_status_col_idx = headers.index(reminder_status_col_name) + 1
        reminder_time_col_idx = headers.index(reminder_time_col_name) + 1
    except ValueError: st.error(f"'{reminder_status_col_name}' or '{reminder_time_col_name}' not in sheet headers."); return False
    row_to_update = find_row_in_gsheet(worksheet, rma, sn, headers)
    if row_to_update != -1:
        reminder_time_str = datetime.combine(reminder_date_obj, datetime.now().time()).strftime("%Y-%m-%d %H:%M:%S")
        updates = [
            {'range': gspread.utils.rowcol_to_a1(row_to_update, reminder_status_col_idx), 'values': [["Yes"]]},
            {'range': gspread.utils.rowcol_to_a1(row_to_update, reminder_time_col_idx), 'values': [[reminder_time_str]]} ]
        if update_gsheet_cells(worksheet, updates):
            st.success(f"Reminder for RMA {rma}, S/N {sn} marked as completed on {reminder_date_obj.strftime('%Y-%m-%d')}."); return True
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
def update_reminder_details_in_gsheet(rma, sn, reminder_date_obj):
    return gsheet_update_wrapper(_update_reminder_in_sheet, rma, sn, reminder_date_obj)
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
    required_cols = ['Estimate Complete Time', 'Estimate Complete', 'Estimate Approved', 'RMA',
                     'Estimate Sent To Email', 'Estimate Sent Time', 
                     'Reminder Completed', 'Reminder Completed Time'] # Added new columns
    if df.empty or not all(col in df.columns for col in required_cols): return pd.DataFrame()
    
    df_copy = df.copy() 
    df_copy['Estimate Complete Time'] = pd.to_datetime(df_copy['Estimate Complete Time'], errors='coerce')
    # Ensure new time columns are also datetime for consistent formatting
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
                
                # Format new time columns for display
                est_sent_time_str = row['Estimate Sent Time'].strftime('%Y-%m-%d %H:%M') if pd.notna(row['Estimate Sent Time']) else 'N/A'
                reminder_time_str = row['Reminder Completed Time'].strftime('%Y-%m-%d %H:%M') if pd.notna(row['Reminder Completed Time']) else 'N/A'

                overdue_items.append({
                    'RMA': rma_value, 
                    'S/N': row.get('S/N', 'N/A'),
                    'Estimate Complete Time': row['Estimate Complete Time'].strftime('%Y-%m-%d'),
                    'Days Pending Approval': (now - row['Estimate Complete Time']).days,
                    'Estimate Sent To Email': row.get('Estimate Sent To Email', 'N/A'), # New
                    'Estimate Sent Time': est_sent_time_str,  # New
                    'Reminder Completed': row.get('Reminder Completed', 'N/A'), # New
                    'Reminder Completed Time': reminder_time_str, # New
                    BC_LINK_COL_NAME: bc_url 
                })
    return pd.DataFrame(overdue_items)

def identify_overdue_for_shipping(df, days_threshold=1):
    required_cols = ['QA Approved Time', 'Estimate Complete', 'Estimate Approved', 'QA Approved', 'Shipped', 'RMA']
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
                overdue_shipping_items.append({ 'RMA': rma_value, 'S/N': row.get('S/N', 'N/A'),
                    'QA Approved Time': row['QA Approved Time'].strftime('%Y-%m-%d'),
                    'Days Pending Shipping': (now - row['QA Approved Time']).days, BC_LINK_COL_NAME: bc_url })
    return pd.DataFrame(overdue_shipping_items)

# --- Main Application ---
st.title("🛠️ Service Process Dashboard") 
st.markdown("Monitor and update service item statuses, including shipping.")

if 'first_load_complete' not in st.session_state: st.session_state.first_load_complete = False
if 'refresh_counter' not in st.session_state: st.session_state.refresh_counter = 0

if 'data_df' not in st.session_state:
    st.session_state.data_df = load_data_from_google_sheet()
    st.session_state.first_load_complete = True 

if st.button("🔄 Refresh Data from Google Sheet"):
    load_data_from_google_sheet.clear(); st.session_state.data_df = load_data_from_google_sheet()
    st.session_state.first_load_complete = False
    st.session_state.refresh_counter += 1 
    st.rerun() 

data_df = st.session_state.data_df

if not data_df.empty:
    st.subheader("📊 Key Metrics")
    display_kpis(data_df.copy()); st.markdown("---")

    st.subheader("⚠️ Overdue Estimates Report (Pending Approval > 3 Days)")
    overdue_estimates_df = identify_overdue_estimates(data_df, days_threshold=3) 
    if not overdue_estimates_df.empty:
        st.warning("The following estimates were completed more than 3 days ago and are still pending approval:")
        # Updated display columns for this report
        overdue_estimates_display_cols = [
            'RMA', 'S/N', 'Estimate Complete Time', 'Days Pending Approval', 
            'Estimate Sent To Email', 'Estimate Sent Time', 
            'Reminder Completed', 'Reminder Completed Time', 
            BC_LINK_COL_NAME
        ]
        # Ensure all display columns exist in the dataframe, add if missing (should be handled by identify_overdue_estimates)
        for col in overdue_estimates_display_cols:
            if col not in overdue_estimates_df.columns:
                overdue_estimates_df[col] = None # Or 'N/A' if preferred for display

        st.dataframe(overdue_estimates_df[overdue_estimates_display_cols], use_container_width=True,
            column_config={ BC_LINK_COL_NAME: st.column_config.LinkColumn(label="Business Central", display_text="Open RMA")},
            column_order=overdue_estimates_display_cols )
    else: st.success("✅ No estimates are currently overdue for approval beyond 3 days.")
    st.markdown("---")

    st.subheader("🚚 Overdue for Shipping Report (QA Approved > 1 Day, Not Shipped)")
    overdue_shipping_df = identify_overdue_for_shipping(data_df, days_threshold=1)
    if not overdue_shipping_df.empty:
        st.error("The following items are QA Approved for more than 1 day and are pending shipment:") 
        overdue_shipping_display_cols = ['RMA', 'S/N', 'QA Approved Time', 'Days Pending Shipping', BC_LINK_COL_NAME]
        if BC_LINK_COL_NAME not in overdue_shipping_df.columns: overdue_shipping_df[BC_LINK_COL_NAME] = None
        st.dataframe(overdue_shipping_df[overdue_shipping_display_cols], use_container_width=True,
            column_config={BC_LINK_COL_NAME: st.column_config.LinkColumn(label="Business Central", display_text="Open RMA")},
            column_order=overdue_shipping_display_cols)
    else: st.success("✅ No items are currently overdue for shipping beyond 1 day.")
    st.markdown("---")

    # --- Sidebar ---
    st.sidebar.header("🔍 Filter Options")
    filtered_df = data_df.copy() 

    for col_name, search_label in [('RMA', "RMA"), ('S/N', "S/N"), ('Part Number', "Part Number"), ('SPC Code', "SPC Code")]:
        if col_name in filtered_df.columns:
            search_term = st.sidebar.text_input(f"Search by {search_label}", key=f"search_{col_name}_{st.session_state.refresh_counter}") 
            if search_term: filtered_df = filtered_df[filtered_df[col_name].astype(str).str.contains(search_term, case=False, na=False)]
    
    status_columns_to_filter = {
        'Estimate Complete': 'Estimate Complete', 'Estimate Approved': 'Estimate Approved',
        'Reminder Completed': 'Reminder Completed', 
        'QA Approved': 'QA Approved', 'Shipped': 'Shipped' 
    }
    for display_name, col_name in status_columns_to_filter.items():
        if col_name in data_df.columns: 
            unique_values = ['All'] + sorted(list(set(val for val in data_df[col_name].astype(str).unique() if val and val.strip() != ''))) 
            if 'N/A' in data_df[col_name].astype(str).unique() and 'N/A' not in unique_values: unique_values.insert(1, "N/A") 
            if 'No' in data_df[col_name].astype(str).unique() and 'No' not in unique_values: unique_values.insert(1, "No") 

            default_index = 0 
            if not st.session_state.first_load_complete:
                if col_name == "Shipped" and "N/A" in unique_values: default_index = unique_values.index("N/A")
                elif col_name == "Reminder Completed" and "No" in unique_values: default_index = unique_values.index("No")
            
            selected_status = st.sidebar.selectbox(f"Filter by {display_name}", unique_values, 
                                                   key=f"select_{col_name}_{st.session_state.refresh_counter}", 
                                                   index=default_index)
            if selected_status != "All": 
                if col_name in filtered_df.columns: 
                    filtered_df = filtered_df[filtered_df[col_name].astype(str) == selected_status]

    st.sidebar.markdown("---"); st.sidebar.subheader("Date Range Filters")
    date_filter_columns_to_filter = {
        'Estimate Complete Time': 'Estimate Complete Time', 'Estimate Approved Time': 'Estimate Approved Time',
        'Estimate Sent Time': 'Estimate Sent Time', 
        'Reminder Completed Time': 'Reminder Completed Time', 
        'QA Approved Time': 'QA Approved Time', 'Shipped Time': 'Shipped Time' 
    }
    for display_name, col_name in date_filter_columns_to_filter.items():
        min_val_for_widget_setup = None; max_val_for_widget_setup = None; can_setup_widget = False
        if col_name in data_df.columns and pd.api.types.is_datetime64_any_dtype(data_df[col_name]):
            original_col_for_widget_params = data_df[col_name].dropna() 
            if not original_col_for_widget_params.empty:
                min_val_for_widget_setup = original_col_for_widget_params.min().date(); max_val_for_widget_setup = original_col_for_widget_params.max().date()
                can_setup_widget = True
        
        if can_setup_widget:
            current_date_range_selection = st.sidebar.date_input(f"Filter by {display_name}", 
                value=[], 
                min_value=min_val_for_widget_setup, max_value=max_val_for_widget_setup, 
                key=f"date_range_{col_name}_{st.session_state.refresh_counter}") 
            
            if current_date_range_selection and len(current_date_range_selection) == 2: 
                if col_name in filtered_df.columns and pd.api.types.is_datetime64_any_dtype(filtered_df[col_name]):
                    start_date_selected, end_date_selected = current_date_range_selection
                    start_datetime_selected = pd.to_datetime(start_date_selected); end_datetime_selected = pd.to_datetime(end_date_selected).replace(hour=23, minute=59, second=59) 
                    condition = ((filtered_df[col_name] >= start_datetime_selected) & (filtered_df[col_name] <= end_datetime_selected) & (filtered_df[col_name].notna()) )
                    filtered_df = filtered_df[condition]
    
    if not st.session_state.first_load_complete: st.session_state.first_load_complete = True

    # --- Sidebar Update Sections ---
    st.sidebar.markdown("---"); st.sidebar.header("📝 Log Estimate Sent")
    if all(c in data_df.columns for c in ['RMA', 'S/N', 'Estimate Approved', 'Estimate Sent To Email']):
        eligible_estimate_sent_df = data_df[
            (data_df['Estimate Approved'].astype(str).str.lower() == 'yes') &
            (data_df['Estimate Sent To Email'].astype(str).str.lower() == 'n/a') ]
        if not eligible_estimate_sent_df.empty:
            options = ["Select item..."] + [f"{rma} - S/N: {sn}" for rma, sn in zip(eligible_estimate_sent_df['RMA'], eligible_estimate_sent_df['S/N'])]
            selected_item_est_sent = st.sidebar.selectbox("Item (Approved, Not Sent)", options, index=0, key=f"est_sent_item_selector_{st.session_state.refresh_counter}")
            if selected_item_est_sent and selected_item_est_sent != "Select item...":
                rma_est_sent, sn_part_est_sent = selected_item_est_sent.split(" - S/N: "); sn_est_sent = sn_part_est_sent.strip()
                sent_to_email = st.sidebar.text_input("Sent To Email Address", key=f"est_sent_email_input_{st.session_state.refresh_counter}")
                sent_date_val = st.sidebar.date_input("Estimate Sent Date", value=date.today(), key=f"est_sent_date_input_{st.session_state.refresh_counter}")
                if st.sidebar.button("Mark Estimate Sent", key=f"mark_est_sent_button_{st.session_state.refresh_counter}"):
                    if rma_est_sent and sn_est_sent and sent_to_email and sent_date_val:
                        if "@" not in sent_to_email or "." not in sent_to_email: st.sidebar.error("Please enter a valid email address.")
                        else:
                            success = update_estimate_sent_details_in_gsheet(rma_est_sent, sn_est_sent, sent_to_email, sent_date_val)
                            if success:
                                load_data_from_google_sheet.clear(); st.session_state.data_df = load_data_from_google_sheet()
                                st.session_state.first_load_complete = False; st.session_state.refresh_counter +=1; st.sidebar.success("Estimate sent details updated!"); st.rerun()
                    else: st.sidebar.warning("Please select item, enter email, and date.")
        else: st.sidebar.info("No items currently eligible for marking as estimate sent.")
    
    st.sidebar.markdown("---"); st.sidebar.header("📞 Log Reminder")
    if all(c in data_df.columns for c in ['RMA', 'S/N', 'Estimate Sent To Email', 'Reminder Completed']):
        eligible_reminder_df = data_df[
            (data_df['Estimate Sent To Email'].astype(str).str.lower() != 'n/a') &
            (data_df['Reminder Completed'].astype(str).str.lower().isin(['no', 'n/a'])) ]
        if not eligible_reminder_df.empty:
            options = ["Select item..."] + [f"{rma} - S/N: {sn}" for rma, sn in zip(eligible_reminder_df['RMA'], eligible_reminder_df['S/N'])]
            selected_item_reminder = st.sidebar.selectbox("Item (Estimate Sent, Reminder Pending)", options, index=0, key=f"reminder_item_selector_{st.session_state.refresh_counter}")
            if selected_item_reminder and selected_item_reminder != "Select item...":
                rma_reminder, sn_part_reminder = selected_item_reminder.split(" - S/N: "); sn_reminder = sn_part_reminder.strip()
                reminder_date_val = st.sidebar.date_input("Reminder Date", value=date.today(), key=f"reminder_date_input_{st.session_state.refresh_counter}")
                if st.sidebar.button("Mark Reminder Completed", key=f"mark_reminder_button_{st.session_state.refresh_counter}"):
                    if rma_reminder and sn_reminder and reminder_date_val:
                        success = update_reminder_details_in_gsheet(rma_reminder, sn_reminder, reminder_date_val)
                        if success:
                            load_data_from_google_sheet.clear(); st.session_state.data_df = load_data_from_google_sheet()
                            st.session_state.first_load_complete = False; st.session_state.refresh_counter +=1; st.sidebar.success("Reminder details updated!"); st.rerun()
                    else: st.sidebar.warning("Please select an item and date.")
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
                            if success:
                                load_data_from_google_sheet.clear(); st.session_state.data_df = load_data_from_google_sheet() 
                                st.session_state.first_load_complete = False; st.session_state.refresh_counter +=1; st.sidebar.success("Update successful! Data refreshed."); st.rerun() 
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
