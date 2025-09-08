import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta # Added timedelta
import gspread 
from oauth2client.service_account import ServiceAccountCredentials 
from io import BytesIO # Added for Excel export
import urllib.parse # Re-added for BC links

# --- Page Configuration ---
st.set_page_config(
    page_title="Service Data Dashboard", 
    page_icon="🚚", 
    layout="wide",
)

# --- Constants for Google Sheets ---
GSHEET_NAME = "Estimate form"
WORKSHEET_INDEX = 1 
CREDS_FILE = "Credentials.json"
EXPECTED_COLUMN_ORDER = [
    "RMA", "SPC Code", "Part Number", "S/N", "Description", 
    "Fault Comments", "Resolution Comments", "Sender", 
    "Estimate Complete Time", "Estimate Complete", 
    "Estimate Approved", "Estimate Approved Time",
    "Estimate Sent To Email", "Estimate Sent Time", # New
    "Reminder Completed", "Reminder Completed Time", # New
    "QA Approved", "QA Approved Time",
    "Shipped", "Shipped Time" 
]
ALL_STATUS_COLUMNS = ["Estimate Complete", "Estimate Approved", "Reminder Completed", "QA Approved", "Shipped"]


# --- Constants for Business Central Link ---
BC_BASE_URL = "https://businesscentral.dynamics.com/7bcfb5b0-27a1-4e18-99d8-ca66570addd8/Production"
BC_COMPANY = "PROD"
BC_PAGE_ID = "70001" # Page for Service Orders / Service Item Worksheet etc.
BC_RMA_FIELD_NAME = "No." # ASSUMPTION: Field name for RMA in Business Central. VERIFY THIS!
BC_LINK_COL_NAME = "View in BC" # Define the column name for BC links

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
            
        headers = all_values[0]
        data_rows = all_values[1:]
        
        # Initialize df with expected columns to ensure structure even if sheet is different
        df = pd.DataFrame(columns=EXPECTED_COLUMN_ORDER)
        temp_df = pd.DataFrame(data_rows, columns=headers) # Load sheet data with its own headers

        for col in EXPECTED_COLUMN_ORDER:
            if col in temp_df.columns:
                df[col] = temp_df[col] # Use data from sheet if column exists
            else: # If expected column is not in sheet, initialize it
                if headers != EXPECTED_COLUMN_ORDER: 
                     st.warning(f"Expected column '{col}' not found in Google Sheet. Initializing as empty.")
                if "Time" in col: df[col] = pd.NaT
                elif col in ALL_STATUS_COLUMNS: df[col] = "No"
                elif col == "Estimate Sent To Email": df[col] = "N/A" # Specific default for this new text field
                else: df[col] = "N/A"
        
        df = df[EXPECTED_COLUMN_ORDER] # Ensure correct order

        string_cols_for_na_fill = ['RMA', 'S/N', 'Part Number', 'SPC Code', 
                                   'Description', 'Fault Comments', 'Resolution Comments', 
                                   'Sender', 'Estimate Sent To Email'] + ALL_STATUS_COLUMNS
        for col in string_cols_for_na_fill:
            if col in df.columns:
                df[col] = df[col].astype(str) 
                df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'N/A') 

        date_cols = [col for col in EXPECTED_COLUMN_ORDER if "Time" in col]
        for col in date_cols:
            if col in df.columns:
                df[col] = df[col].replace('N/A', None) 
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        return df

    except FileNotFoundError:
        st.error(f"Error: Credentials file '{creds_file}' not found. Please ensure it's in the correct path.")
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"Error: Google Sheet '{sheet_name}' not found. Please check the name and permissions.")
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Error: Worksheet with index {worksheet_index} not found in '{sheet_name}'.")
    except Exception as e:
        st.error(f"An error occurred while loading data from Google Sheets: {e}")
    return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) 

def find_row_in_gsheet(worksheet, rma_to_find, sn_to_find, headers):
    """Helper to find a row index based on RMA and S/N."""
    try:
        rma_col_index = headers.index("RMA") + 1
        sn_col_index = headers.index("S/N") + 1
    except ValueError:
        st.error("RMA or S/N column not found in sheet headers during row search.")
        return -1

    all_data_with_headers = worksheet.get_all_values() # Potentially slow for very large sheets
    for i, row_values in enumerate(all_data_with_headers[1:], start=2): 
        rma_val = row_values[rma_col_index - 1] if len(row_values) >= rma_col_index else None
        sn_val = row_values[sn_col_index - 1] if len(row_values) >= sn_col_index else None
        if rma_val == rma_to_find and sn_val == sn_to_find:
            return i
    return -1

def update_gsheet_cells(worksheet, updates_list):
    """Helper to perform batch updates on a worksheet."""
    try:
        worksheet.batch_update(updates_list)
        return True
    except Exception as e:
        st.error(f"An error occurred during Google Sheet batch update: {e}")
        return False

def update_estimate_sent_details_in_gsheet(rma, sn, sent_to_email, sent_date_obj):
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open(GSHEET_NAME)
        worksheet = spreadsheet.get_worksheet(WORKSHEET_INDEX)
        headers = worksheet.row_values(1)
        if not headers: return False

        sent_time_col_index = headers.index("Estimate Sent Time") + 1
        sent_email_col_index = headers.index("Estimate Sent To Email") + 1
        
        row_to_update = find_row_in_gsheet(worksheet, rma, sn, headers)
        if row_to_update != -1:
            sent_time_str = datetime.combine(sent_date_obj, datetime.now().time()).strftime("%Y-%m-%d %H:%M:%S")
            updates = [
                {'range': gspread.utils.rowcol_to_a1(row_to_update, sent_email_col_index), 'values': [[sent_to_email]]},
                {'range': gspread.utils.rowcol_to_a1(row_to_update, sent_time_col_index), 'values': [[sent_time_str]]}
            ]
            if update_gsheet_cells(worksheet, updates):
                st.success(f"Estimate for RMA {rma}, S/N {sn} marked as sent to {sent_to_email} on {sent_date_obj.strftime('%Y-%m-%d')}.")
                return True
        else:
            st.error(f"Record for RMA {rma}, S/N {sn} not found for estimate sent update.")
        return False
    except Exception as e:
        st.error(f"Error updating estimate sent details: {e}")
        return Fa