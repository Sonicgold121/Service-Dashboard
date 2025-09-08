import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta # Added timedelta
import gspread 
from oauth2client.service_account import ServiceAccountCredentials 
from io import BytesIO # Added for Excel export

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
    "QA Approved", "QA Approved Time",
    "Shipped", "Shipped Time" 
]


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
        
        if headers != EXPECTED_COLUMN_ORDER:
            st.warning(f"Google Sheet headers {headers} do not perfectly match expected {EXPECTED_COLUMN_ORDER}. Using sheet headers for DataFrame creation, then aligning.")
        
        df = pd.DataFrame(data_rows, columns=headers) 
        
        for col in EXPECTED_COLUMN_ORDER:
            if col not in df.columns:
                if "Time" in col:
                    df[col] = pd.NaT
                elif col in ["Estimate Complete", "Estimate Approved", "QA Approved", "Shipped"]:
                    df[col] = "No" 
                else:
                    df[col] = "N/A" 

        cols_to_use = [col for col in EXPECTED_COLUMN_ORDER] 
        df = df[cols_to_use]

        string_cols_for_na_fill = ['RMA', 'S/N', 'Part Number', 'SPC Code', 
                                   'Description', 'Fault Comments', 'Resolution Comments', 
                                   'Sender', 'Estimate Complete', 'Estimate Approved', 'QA Approved', 'Shipped'] 
        for col in string_cols_for_na_fill:
            if col in df.columns:
                df[col] = df[col].astype(str) 
                df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'N/A') 

        date_cols = ['Estimate Complete Time', 'Estimate Approved Time', 'QA Approved Time', 'Shipped Time'] 
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


def update_shipped_status_in_gsheet(rma_to_update, sn_to_update, shipped_date_obj):
    """Updates the Google Sheet for the given RMA and S/N with shipped status and date."""
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

        try:
            rma_col_header = "RMA"; sn_col_header = "S/N"; shipped_status_header = "Shipped"; shipped_time_header = "Shipped Time"
            if not all(h in headers for h in [rma_col_header, sn_col_header, shipped_status_header, shipped_time_header]):
                 missing = [h for h in [rma_col_header, sn_col_header, shipped_status_header, shipped_time_header] if h not in headers]
                 st.error(f"Required columns {missing} not found in Google Sheet headers. Update failed.")
                 return False
            rma_col_index = headers.index(rma_col_header) + 1; sn_col_index = headers.index(sn_col_header) + 1
            shipped_status_col_index = headers.index(shipped_status_header) + 1; shipped_time_col_index = headers.index(shipped_time_header) + 1
        except ValueError as ve: 
            st.error(f"Error finding column index in Google Sheet headers: {ve}. Update failed.")
            return False

        all_data_with_headers = worksheet.get_all_values()
        row_to_update = -1
        
        for i, row_values in enumerate(all_data_with_headers[1:], start=2): 
            rma_val = row_values[rma_col_index - 1] if len(row_values) >= rma_col_index else None
            sn_val = row_values[sn_col_index - 1] if len(row_values) >= sn_col_index else None
            if rma_val == rma_to_update and sn_val == sn_to_update:
                row_to_update = i; break
        
        if row_to_update != -1:
            shipped_date_str = shipped_date_obj.strftime("%Y-%m-%d %H:%M:%S") 
            update_payload = [
                {'range': gspread.utils.rowcol_to_a1(row_to_update, shipped_status_col_index), 'values': [["Yes"]]},
                {'range': gspread.utils.rowcol_to_a1(row_to_update, shipped_time_col_index), 'values': [[shipped_date_str]]}
            ]
            worksheet.batch_update(update_payload)
            st.success(f"Successfully marked RMA {rma_to_update}, S/N {sn_to_update} as shipped on {shipped_date_obj.strftime('%Y-%m-%d')}.")
            return True
        else:
            st.error(f"Record with RMA {rma_to_update} and S/N {sn_to_update} not found in Google Sheet. Update failed.")
            return False
    except Exception as e:
        st.error(f"An error occurred while updating Google Sheet: {e}")
        return False

def display_kpis(df):
    """Displays Key Performance Indicators."""
    if df.empty: return
    total_records = len(df)
    estimate_complete_count = df[df['Estimate Complete'].astype(str).str.lower() == 'yes'].shape[0] if 'Estimate Complete' in df.columns else 0
    estimate_approved_count = df[df['Estimate Approved'].astype(str).str.lower() == 'yes'].shape[0] if 'Estimate Approved' in df.columns else 0
    qa_approved_count = df[df['QA Approved'].astype(str).str.lower() == 'yes'].shape[0] if 'QA Approved' in df.columns else 0
    shipped_count = df[df['Shipped'].astype(str).str.lower() == 'yes'].shape[0] if 'Shipped' in df.columns else 0
    cols = st.columns(5) 
    cols[0].metric("Total Records", total_records); cols[1].metric("Estimates Complete", estimate_complete_count)
    cols[2].metric("Estimates Approved", estimate_approved_count); cols[3].metric("QA Approved", qa_approved_count)
    cols[4].metric("Units Shipped", shipped_count) 

# --- NEW: Function to identify overdue estimates ---
def identify_overdue_estimates(df, days_threshold=3):
    """Identifies estimates that are complete but not yet approved for more than X days."""
    if df.empty or 'Estimate Complete Time' not in df.columns or \
       'Estimate Complete' not in df.columns or 'Estimate Approved' not in df.columns:
        return pd.DataFrame()

    # Ensure 'Estimate Complete Time' is datetime
    df_copy = df.copy() # Work on a copy
    df_copy['Estimate Complete Time'] = pd.to_datetime(df_copy['Estimate Complete Time'], errors='coerce')

    now = datetime.now()
    overdue_items = []

    for index, row in df_copy.iterrows():
        is_complete = str(row.get('Estimate Complete', 'N/A')).lower() == 'yes'
        is_not_approved = str(row.get('Estimate Approved', 'N/A')).lower() in ['no', 'n/a']
        complete_time = row['Estimate Complete Time']

        if is_complete and is_not_approved and pd.notna(complete_time):
            days_passed = (now - complete_time).days
            if days_passed > days_threshold:
                overdue_info = {
                    'RMA': row.get('RMA', 'N/A'),
                    'S/N': row.get('S/N', 'N/A'),
                    'Estimate Complete Time': complete_time.strftime('%Y-%m-%d') if pd.notna(complete_time) else 'N/A',
                    'Days Pending Approval': days_passed
                }
                overdue_items.append(overdue_info)
    
    return pd.DataFrame(overdue_items)

# --- Main Application ---
st.title("🛠️ Service Process Dashboard") 
st.markdown("Monitor and update service item statuses, including shipping.")

if 'data_df' not in st.session_state:
    st.session_state.data_df = load_data_from_google_sheet()

if st.button("🔄 Refresh Data from Google Sheet"):
    load_data_from_google_sheet.clear() 
    st.session_state.data_df = load_data_from_google_sheet() 
    st.rerun() 

data_df = st.session_state.data_df

if not data_df.empty:
    st.subheader("📊 Key Metrics")
    display_kpis(data_df.copy()) 
    st.markdown("---")

    # --- NEW: Overdue Estimates Report Section ---
    st.subheader("⚠️ Overdue Estimates Report (Pending Approval > 3 Days)")
    overdue_df = identify_overdue_estimates(data_df, days_threshold=3)
    if not overdue_df.empty:
        st.warning("The following estimates were completed more than 3 days ago and are still pending approval:")
        st.dataframe(overdue_df, use_container_width=True)
    else:
        st.success("✅ No estimates are currently overdue for approval beyond 3 days.")
    st.markdown("---")
    # --- END NEW SECTION ---


    st.sidebar.header("🔍 Filter Options")
    filtered_df = data_df.copy()

    for col_name, search_label in [('RMA', "RMA"), ('S/N', "S/N"), 
                                   ('Part Number', "Part Number"), ('SPC Code', "SPC Code")]:
        if col_name in filtered_df.columns:
            search_term = st.sidebar.text_input(f"Search by {search_label}", key=f"search_{col_name}")
            if search_term:
                filtered_df = filtered_df[filtered_df[col_name].astype(str).str.contains(search_term, case=False, na=False)]
    
    status_columns_to_filter = {
        'Estimate Complete': 'Estimate Complete', 'Estimate Approved': 'Estimate Approved',
        'QA Approved': 'QA Approved', 'Shipped': 'Shipped' 
    }
    for display_name, col_name in status_columns_to_filter.items():
        if col_name in filtered_df.columns:
            unique_values = ['All'] + sorted(filtered_df[col_name].astype(str).unique().tolist())
            selected_status = st.sidebar.selectbox(f"Filter by {display_name}", unique_values, key=f"select_{col_name}")
            if selected_status != "All":
                filtered_df = filtered_df[filtered_df[col_name].astype(str) == selected_status]

    st.sidebar.markdown("---")
    st.sidebar.subheader("Date Range Filters")
    date_filter_columns_to_filter = {
        'Estimate Complete Time': 'Estimate Complete Time', 'Estimate Approved Time': 'Estimate Approved Time',
        'QA Approved Time': 'QA Approved Time', 'Shipped Time': 'Shipped Time' 
    }

    for display_name, col_name in date_filter_columns_to_filter.items():
        if col_name in filtered_df.columns and pd.api.types.is_datetime64_any_dtype(filtered_df[col_name]):
            temp_col_dt = filtered_df[col_name] 
            if temp_col_dt.notna().any(): 
                min_date_val = temp_col_dt.min(); max_date_val = temp_col_dt.max() 
                if pd.isna(min_date_val) or pd.isna(max_date_val): continue
                min_date_dt = min_date_val.date(); max_date_dt = max_date_val.date()
                try:
                    date_range = st.sidebar.date_input(
                        f"Filter by {display_name}", value=(min_date_dt, max_date_dt),
                        min_value=min_date_dt, max_value=max_date_dt, key=f"date_range_{col_name}"
                    )
                    if date_range and len(date_range) == 2:
                        start_date_dt, end_date_dt = date_range
                        start_datetime = pd.to_datetime(start_date_dt)
                        end_datetime = pd.to_datetime(end_date_dt).replace(hour=23, minute=59, second=59)
                        filtered_df = filtered_df[
                            (filtered_df[col_name] >= start_datetime) & 
                            (filtered_df[col_name] <= end_datetime) &
                            (filtered_df[col_name].notna()) ]
                except Exception as e: st.sidebar.error(f"Error with date filter for {display_name}: {e}")

    st.sidebar.markdown("---")
    st.sidebar.header("📦 Update Shipped Status")
    if 'RMA' in data_df.columns and 'S/N' in data_df.columns and 'Shipped' in data_df.columns:
        unshipped_items_df = data_df[data_df['Shipped'].astype(str).str.lower().isin(['no', 'n/a'])]
        if not unshipped_items_df.empty:
            unshipped_options = ["Select an item..."] + [f"{rma} - S/N: {sn}" for rma, sn in zip(unshipped_items_df['RMA'], unshipped_items_df['S/N'])]
            selected_item_str = st.sidebar.selectbox("Select Item to Mark as Shipped (RMA - S/N)", options=unshipped_options, index=0, key="shipped_item_selector")
            if selected_item_str and selected_item_str != "Select an item...":
                try:
                    rma_to_update, sn_part = selected_item_str.split(" - S/N: "); sn_to_update = sn_part.strip()
                    shipped_date_val = st.sidebar.date_input("Shipped Date", value=date.today(), key="shipped_date_input") 
                    if st.sidebar.button("Mark as Shipped", key="mark_shipped_button"):
                        if rma_to_update and sn_to_update and shipped_date_val:
                            success = update_shipped_status_in_gsheet(rma_to_update, sn_to_update, shipped_date_val)
                            if success:
                                load_data_from_google_sheet.clear(); st.session_state.data_df = load_data_from_google_sheet() 
                                st.sidebar.success("Update successful! Data refreshed."); st.rerun() 
                            else: st.sidebar.error("Update failed. Check logs or details above.")
                        else: st.sidebar.warning("Please select an item and a valid shipped date.")
                except ValueError: st.sidebar.error("Invalid item format selected. Please re-select.")
        elif not data_df.empty : st.sidebar.info("All available items are marked as shipped.")

    st.subheader("Filtered Data View")
    st.markdown(f"Displaying **{len(filtered_df)}** records out of **{len(data_df) if not data_df.empty else 0}** total records.")
    if not filtered_df.empty:
        display_cols = [col for col in EXPECTED_COLUMN_ORDER if col in filtered_df.columns]
        st.dataframe(filtered_df[display_cols].astype(str), use_container_width=True)
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
