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
    "QA Approved", "QA Approved Time",
    "Shipped", "Shipped Time" 
]

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
                if headers != EXPECTED_COLUMN_ORDER: # Only warn if headers were actually different
                     st.warning(f"Expected column '{col}' not found in Google Sheet. Initializing as empty.")
                if "Time" in col: df[col] = pd.NaT
                elif col in ["Estimate Complete", "Estimate Approved", "QA Approved", "Shipped"]: df[col] = "No"
                else: df[col] = "N/A"
        
        # Ensure correct order
        df = df[EXPECTED_COLUMN_ORDER]

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

def identify_overdue_estimates(df, days_threshold=3):
    """Identifies estimates that are complete but not yet approved for more than X days."""
    if df.empty or 'Estimate Complete Time' not in df.columns or \
       'Estimate Complete' not in df.columns or 'Estimate Approved' not in df.columns:
        return pd.DataFrame()

    df_copy = df.copy() 
    df_copy['Estimate Complete Time'] = pd.to_datetime(df_copy['Estimate Complete Time'], errors='coerce')
    now = datetime.now()
    overdue_items = []

    for index, row in df_copy.iterrows():
        is_complete = str(row.get('Estimate Complete', 'N/A')).lower() == 'yes'
        is_not_approved = str(row.get('Estimate Approved', 'N/A')).lower() in ['no', 'n/a']
        complete_time = row['Estimate Complete Time']
        rma_value = str(row.get('RMA', 'N/A'))

        if is_complete and is_not_approved and pd.notna(complete_time):
            days_passed = (now - complete_time).days
            if days_passed > days_threshold:
                bc_url = None
                if rma_value != 'N/A' and rma_value.strip() != "":
                    bc_url = f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(rma_value)}%27"
                
                overdue_info = {
                    'RMA': rma_value, 
                    'S/N': row.get('S/N', 'N/A'),
                    'Estimate Complete Time': complete_time.strftime('%Y-%m-%d') if pd.notna(complete_time) else 'N/A',
                    'Days Pending Approval': days_passed,
                    BC_LINK_COL_NAME: bc_url # Add the BC link URL
                }
                overdue_items.append(overdue_info)
    
    return pd.DataFrame(overdue_items)

# --- NEW: Function to identify items overdue for shipping ---
def identify_overdue_for_shipping(df, days_threshold=1):
    """Identifies items that are QA Approved but not Shipped for more than X days."""
    if df.empty or 'QA Approved Time' not in df.columns or \
       'Estimate Complete' not in df.columns or 'Estimate Approved' not in df.columns or \
       'QA Approved' not in df.columns or 'Shipped' not in df.columns:
        return pd.DataFrame()

    df_copy = df.copy()
    df_copy['QA Approved Time'] = pd.to_datetime(df_copy['QA Approved Time'], errors='coerce')
    now = datetime.now()
    overdue_shipping_items = []

    for index, row in df_copy.iterrows():
        is_estimate_complete = str(row.get('Estimate Complete', 'N/A')).lower() == 'yes'
        is_estimate_approved = str(row.get('Estimate Approved', 'N/A')).lower() == 'yes'
        is_qa_approved = str(row.get('QA Approved', 'N/A')).lower() == 'yes'
        is_not_shipped = str(row.get('Shipped', 'N/A')).lower() in ['no', 'n/a']
        qa_approved_time = row['QA Approved Time']
        rma_value = str(row.get('RMA', 'N/A'))

        if is_estimate_complete and is_estimate_approved and is_qa_approved and \
           is_not_shipped and pd.notna(qa_approved_time):
            days_passed_shipping = (now - qa_approved_time).days
            if days_passed_shipping > days_threshold:
                bc_url = None
                if rma_value != 'N/A' and rma_value.strip() != "":
                    bc_url = f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(rma_value)}%27"

                overdue_shipping_info = {
                    'RMA': rma_value,
                    'S/N': row.get('S/N', 'N/A'),
                    'QA Approved Time': qa_approved_time.strftime('%Y-%m-%d') if pd.notna(qa_approved_time) else 'N/A',
                    'Days Pending Shipping': days_passed_shipping,
                    BC_LINK_COL_NAME: bc_url
                }
                overdue_shipping_items.append(overdue_shipping_info)
                
    return pd.DataFrame(overdue_shipping_items)


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

    # --- Overdue Estimates Report Section ---
    st.subheader("⚠️ Overdue Estimates Report (Pending Approval > 3 Days)")
    overdue_estimates_df = identify_overdue_estimates(data_df, days_threshold=3) # Renamed df variable
    if not overdue_estimates_df.empty:
        st.warning("The following estimates were completed more than 3 days ago and are still pending approval:")
        overdue_estimates_display_cols = ['RMA', 'S/N', 'Estimate Complete Time', 'Days Pending Approval', BC_LINK_COL_NAME]
        if BC_LINK_COL_NAME not in overdue_estimates_df.columns:
            overdue_estimates_df[BC_LINK_COL_NAME] = None 
            
        st.dataframe(
            overdue_estimates_df[overdue_estimates_display_cols], 
            use_container_width=True,
            column_config={
                BC_LINK_COL_NAME: st.column_config.LinkColumn(
                    label="Business Central", display_text="Open RMA",
                )
            },
            column_order=overdue_estimates_display_cols 
        )
    else:
        st.success("✅ No estimates are currently overdue for approval beyond 3 days.")
    st.markdown("---")

    # --- NEW: Overdue for Shipping Report Section ---
    st.subheader("🚚 Overdue for Shipping Report (QA Approved > 1 Day, Not Shipped)")
    overdue_shipping_df = identify_overdue_for_shipping(data_df, days_threshold=1)
    if not overdue_shipping_df.empty:
        st.error("The following items are QA Approved for more than 1 day and are pending shipment:") # Using st.error for higher visibility
        overdue_shipping_display_cols = ['RMA', 'S/N', 'QA Approved Time', 'Days Pending Shipping', BC_LINK_COL_NAME]
        if BC_LINK_COL_NAME not in overdue_shipping_df.columns:
             overdue_shipping_df[BC_LINK_COL_NAME] = None
        
        st.dataframe(
            overdue_shipping_df[overdue_shipping_display_cols],
            use_container_width=True,
            column_config={
                BC_LINK_COL_NAME: st.column_config.LinkColumn(
                    label="Business Central", display_text="Open RMA",
                )
            },
            column_order=overdue_shipping_display_cols
        )
    else:
        st.success("✅ No items are currently overdue for shipping beyond 1 day.")
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
        min_val_for_widget_setup = None; max_val_for_widget_setup = None; can_setup_widget = False
        if col_name in data_df.columns and pd.api.types.is_datetime64_any_dtype(data_df[col_name]):
            original_col_for_widget_params = data_df[col_name].dropna() 
            if not original_col_for_widget_params.empty:
                min_val_for_widget_setup = original_col_for_widget_params.min().date()
                max_val_for_widget_setup = original_col_for_widget_params.max().date()
                can_setup_widget = True
        if can_setup_widget:
            current_date_range_selection = st.sidebar.date_input(
                f"Filter by {display_name}", value=(min_val_for_widget_setup, max_val_for_widget_setup), 
                min_value=min_val_for_widget_setup, max_value=max_val_for_widget_setup, key=f"date_range_{col_name}"
            )
            if col_name in filtered_df.columns and pd.api.types.is_datetime64_any_dtype(filtered_df[col_name]):
                if current_date_range_selection and len(current_date_range_selection) == 2:
                    start_date_selected, end_date_selected = current_date_range_selection
                    start_datetime_selected = pd.to_datetime(start_date_selected) 
                    end_datetime_selected = pd.to_datetime(end_date_selected).replace(hour=23, minute=59, second=59) 
                    condition = ((filtered_df[col_name] >= start_datetime_selected) & (filtered_df[col_name] <= end_datetime_selected) & (filtered_df[col_name].notna()) )
                    filtered_df = filtered_df[condition]

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
        df_for_display = filtered_df.copy()
        df_for_display[BC_LINK_COL_NAME] = df_for_display.apply(
            lambda row: f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(str(row['RMA']))}%27"
            if pd.notna(row['RMA']) and str(row['RMA']).strip() != 'N/A' and str(row['RMA']).strip() != "" else None,
            axis=1
        )
        
        display_cols_order = EXPECTED_COLUMN_ORDER[:] 
        if 'RMA' in display_cols_order:
            rma_index = display_cols_order.index('RMA')
            display_cols_order.insert(rma_index + 1, BC_LINK_COL_NAME)
        else: 
            display_cols_order.append(BC_LINK_COL_NAME)

        final_display_columns = [col for col in display_cols_order if col in df_for_display.columns]

        st.dataframe(
            df_for_display[final_display_columns], 
            use_container_width=True,
            column_config={
                BC_LINK_COL_NAME: st.column_config.LinkColumn(
                    label="Business Central", 
                    display_text="Open RMA", 
                )
            }
        )
    else: 
        st.warning("No data matches the current filter criteria or no data loaded.")

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
