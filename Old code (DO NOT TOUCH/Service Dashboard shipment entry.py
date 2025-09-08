import streamlit as st
import pandas as pd
from datetime import datetime, date
import gspread 
from oauth2client.service_account import ServiceAccountCredentials 

# --- Page Configuration ---
st.set_page_config(
    page_title="Service Data Dashboard", # Changed title slightly
    page_icon="🚚", # Changed icon
    layout="wide",
)

# --- Constants for Google Sheets ---
GSHEET_NAME = "Estimate form"
WORKSHEET_INDEX = 1 # Second sheet
CREDS_FILE = "Credentials.json"
# Define the expected column order, including the new Shipped fields
# This should match the order in your Google Sheet and your email processing script
EXPECTED_COLUMN_ORDER = [
    "RMA", "SPC Code", "Part Number", "S/N", "Description", 
    "Fault Comments", "Resolution Comments", "Sender", 
    "Estimate Complete Time", "Estimate Complete", 
    "Estimate Approved", "Estimate Approved Time",
    "QA Approved", "QA Approved Time",
    "Shipped", "Shipped Time" 
]


# --- Helper Functions ---
@st.cache_data(ttl=300) # Cache data for 5 minutes
def load_data_from_google_sheet(
    sheet_name=GSHEET_NAME, 
    worksheet_index=WORKSHEET_INDEX, 
    creds_file=CREDS_FILE
):
    """Loads data from the specified Google Sheet."""
    st.write(f"Attempting to load data from Google Sheet: {sheet_name}, Worksheet Index: {worksheet_index}")
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
            return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) # Return empty df with expected columns
            
        headers = all_values[0]
        data_rows = all_values[1:]
        
        # Ensure headers match EXPECTED_COLUMN_ORDER or handle discrepancies
        if headers != EXPECTED_COLUMN_ORDER:
            st.warning(f"Google Sheet headers {headers} do not perfectly match expected {EXPECTED_COLUMN_ORDER}. Proceeding with sheet headers.")
            # For robustness, you might want to map sheet headers to expected headers if they differ but mean the same thing.
            # For now, we'll use the sheet's headers.
        
        df = pd.DataFrame(data_rows, columns=headers) # Use actual headers from sheet
        
        # Ensure all EXPECTED_COLUMN_ORDER columns exist in df, add if missing
        for col in EXPECTED_COLUMN_ORDER:
            if col not in df.columns:
                st.info(f"Column '{col}' from EXPECTED_COLUMN_ORDER not found in sheet. Adding it as empty.")
                if "Time" in col:
                    df[col] = pd.NaT
                elif col in ["Estimate Complete", "Estimate Approved", "QA Approved", "Shipped"]:
                    df[col] = "No" # Default status
                else:
                    df[col] = "N/A" # Default for other text

        # Reorder df columns to match EXPECTED_COLUMN_ORDER for consistency
        # Filter EXPECTED_COLUMN_ORDER to only include columns actually present in df to avoid errors if sheet is very different
        cols_to_use = [col for col in EXPECTED_COLUMN_ORDER if col in df.columns]
        df = df[cols_to_use]

        # --- Data Cleaning and Type Conversion ---
        string_cols_for_na_fill = ['RMA', 'S/N', 'Part Number', 'SPC Code', 
                                   'Description', 'Fault Comments', 'Resolution Comments', 
                                   'Sender', 'Estimate Complete', 'Estimate Approved', 'QA Approved', 'Shipped'] # Added Shipped
        for col in string_cols_for_na_fill:
            if col in df.columns:
                df[col] = df[col].astype(str) 
                df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'N/A') # Replace NaT string too

        date_cols = ['Estimate Complete Time', 'Estimate Approved Time', 'QA Approved Time', 'Shipped Time'] # Added Shipped Time
        for col in date_cols:
            if col in df.columns:
                df[col] = df[col].replace('N/A', None) # Replace our "N/A" with None for pd.to_datetime
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        st.success("Data loaded successfully from Google Sheets.")
        return df

    except FileNotFoundError:
        st.error(f"Error: Credentials file '{creds_file}' not found. Please ensure it's in the correct path.")
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"Error: Google Sheet '{sheet_name}' not found. Please check the name and permissions.")
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Error: Worksheet with index {worksheet_index} not found in '{sheet_name}'.")
    except Exception as e:
        st.error(f"An error occurred while loading data from Google Sheets: {e}")
        st.error(f"Details: {type(e).__name__} - {str(e)}")
    return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) # Return empty df with expected columns on error


def update_shipped_status_in_gsheet(rma_to_update, sn_to_update, shipped_date_obj):
    """Updates the Google Sheet for the given RMA and S/N with shipped status and date."""
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open(GSHEET_NAME)
        worksheet = spreadsheet.get_worksheet(WORKSHEET_INDEX)

        headers = worksheet.row_values(1) # Get header row
        if not headers:
            st.error("Could not read headers from Google Sheet. Update failed.")
            return False

        try:
            rma_col_index = headers.index("RMA") + 1  # 1-based index
            sn_col_index = headers.index("S/N") + 1
            shipped_status_col_index = headers.index("Shipped") + 1
            shipped_time_col_index = headers.index("Shipped Time") + 1
        except ValueError as ve:
            st.error(f"One or more required columns (RMA, S/N, Shipped, Shipped Time) not found in Google Sheet headers: {ve}. Update failed.")
            return False

        # Find the row: Iterate through rows to find the match
        # Using get_all_records can be slow for large sheets if we only need indices.
        # A more efficient way for larger sheets might be to find cells.
        # For simplicity with moderate sized sheets:
        all_data_with_headers = worksheet.get_all_values()
        row_to_update = -1
        
        # Iterate from the first data row (index 1, as 0 is header)
        for i, row_values in enumerate(all_data_with_headers[1:], start=2): # start=2 for 1-based sheet row index
            # Ensure row_values has enough elements to prevent IndexError
            rma_val = row_values[rma_col_index - 1] if len(row_values) >= rma_col_index else None
            sn_val = row_values[sn_col_index - 1] if len(row_values) >= sn_col_index else None
            
            if rma_val == rma_to_update and sn_val == sn_to_update:
                row_to_update = i
                break
        
        if row_to_update != -1:
            shipped_date_str = shipped_date_obj.strftime("%Y-%m-%d %H:%M:%S") # Store with time for consistency
            worksheet.update_cell(row_to_update, shipped_status_col_index, "Yes")
            worksheet.update_cell(row_to_update, shipped_time_col_index, shipped_date_str)
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
    if df.empty:
        st.info("No data available for KPIs.")
        return

    total_records = len(df)
    estimate_complete_count = df[df['Estimate Complete'].astype(str).str.lower() == 'yes'].shape[0] if 'Estimate Complete' in df.columns else 0
    estimate_approved_count = df[df['Estimate Approved'].astype(str).str.lower() == 'yes'].shape[0] if 'Estimate Approved' in df.columns else 0
    qa_approved_count = df[df['QA Approved'].astype(str).str.lower() == 'yes'].shape[0] if 'QA Approved' in df.columns else 0
    shipped_count = df[df['Shipped'].astype(str).str.lower() == 'yes'].shape[0] if 'Shipped' in df.columns else 0


    cols = st.columns(5) # Adjusted for new KPI
    cols[0].metric("Total Records", total_records)
    cols[1].metric("Estimates Complete", estimate_complete_count)
    cols[2].metric("Estimates Approved", estimate_approved_count)
    cols[3].metric("QA Approved", qa_approved_count)
    cols[4].metric("Units Shipped", shipped_count) # New KPI

# --- Main Application ---
st.title("🛠️ Service Process Dashboard") # Updated title
st.markdown("Monitor and update service item statuses, including shipping.")

# --- Load Data ---
# Initialize session state for data_df if it doesn't exist
if 'data_df' not in st.session_state:
    st.session_state.data_df = load_data_from_google_sheet()

if st.button("🔄 Refresh Data from Google Sheet"):
    st.session_state.data_df = load_data_from_google_sheet() # Reload data
    st.rerun() # Rerun to update display immediately

data_df = st.session_state.data_df


if not data_df.empty:
    st.subheader("📊 Key Metrics")
    display_kpis(data_df.copy()) 
    st.markdown("---")

    # --- Sidebar for Filters ---
    st.sidebar.header("🔍 Filter Options")
    filtered_df = data_df.copy()

    # Text search
    for col_name, search_label in [('RMA', "RMA"), ('S/N', "S/N"), 
                                   ('Part Number', "Part Number"), ('SPC Code', "SPC Code")]:
        if col_name in filtered_df.columns:
            search_term = st.sidebar.text_input(f"Search by {search_label}", key=f"search_{col_name}")
            if search_term:
                filtered_df = filtered_df[filtered_df[col_name].astype(str).str.contains(search_term, case=False, na=False)]
    
    # Dropdown filters for status columns
    status_columns_to_filter = {
        'Estimate Complete': 'Estimate Complete',
        'Estimate Approved': 'Estimate Approved',
        'QA Approved': 'QA Approved',
        'Shipped': 'Shipped' # Added Shipped filter
    }
    for display_name, col_name in status_columns_to_filter.items():
        if col_name in filtered_df.columns:
            unique_values = ['All'] + sorted(filtered_df[col_name].astype(str).unique().tolist())
            selected_status = st.sidebar.selectbox(f"Filter by {display_name}", unique_values, key=f"select_{col_name}")
            if selected_status != "All":
                filtered_df = filtered_df[filtered_df[col_name].astype(str) == selected_status]

    # Date range filters
    st.sidebar.markdown("---")
    st.sidebar.subheader("Date Range Filters")
    date_filter_columns_to_filter = {
        'Estimate Complete Time': 'Estimate Complete Time',
        'Estimate Approved Time': 'Estimate Approved Time',
        'QA Approved Time': 'QA Approved Time',
        'Shipped Time': 'Shipped Time' # Added Shipped Time filter
    }

    for display_name, col_name in date_filter_columns_to_filter.items():
        if col_name in filtered_df.columns and pd.api.types.is_datetime64_any_dtype(filtered_df[col_name]):
            temp_col_dt = filtered_df[col_name] 
            if temp_col_dt.notna().any(): 
                min_date_val = temp_col_dt.min() 
                max_date_val = temp_col_dt.max() 
                if pd.isna(min_date_val) or pd.isna(max_date_val): 
                    st.sidebar.warning(f"Not enough valid date data in '{display_name}' for range filter.")
                    continue
                
                min_date_dt = min_date_val.date() 
                max_date_dt = max_date_val.date()
                
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
                            (filtered_df[col_name].notna()) 
                        ]
                except Exception as e:
                    st.sidebar.error(f"Error with date filter for {display_name}: {e}")
            else:
                st.sidebar.info(f"No valid date data in '{display_name}' for filtering.")
        elif col_name in filtered_df.columns: # Column exists but not datetime
             st.sidebar.info(f"Column '{display_name}' is not in a recognized date format for filtering.")


    # --- Update Shipped Status Section ---
    st.sidebar.markdown("---")
    st.sidebar.header("📦 Update Shipped Status")
    
    # Create a list of RMA-S/N for selection, preferably for items not yet shipped
    # Ensure 'RMA' and 'S/N' columns exist
    if 'RMA' in data_df.columns and 'S/N' in data_df.columns and 'Shipped' in data_df.columns:
        # Filter for items not yet shipped or where Shipped is 'N/A' or 'No'
        unshipped_items_df = data_df[data_df['Shipped'].astype(str).str.lower().isin(['no', 'n/a'])]
        
        if not unshipped_items_df.empty:
            # Create a unique identifier for the selectbox
            unshipped_options = ["Select an item..."] + [
                f"{rma} - S/N: {sn}" for rma, sn in zip(unshipped_items_df['RMA'], unshipped_items_df['S/N'])
            ]
            selected_item_str = st.sidebar.selectbox(
                "Select Item to Mark as Shipped (RMA - S/N)",
                options=unshipped_options,
                index=0, # Default to "Select an item..."
                key="shipped_item_selector"
            )

            if selected_item_str and selected_item_str != "Select an item...":
                try:
                    rma_to_update, sn_part = selected_item_str.split(" - S/N: ")
                    sn_to_update = sn_part.strip()

                    shipped_date = st.sidebar.date_input("Shipped Date", value=date.today(), key="shipped_date_input")
                    
                    if st.sidebar.button("Mark as Shipped", key="mark_shipped_button"):
                        if rma_to_update and sn_to_update and shipped_date:
                            success = update_shipped_status_in_gsheet(rma_to_update, sn_to_update, shipped_date)
                            if success:
                                st.session_state.data_df = load_data_from_google_sheet() # Reload data
                                st.sidebar.success("Update successful! Refreshing data...") # Show in sidebar
                                st.rerun() # Rerun to update display
                            else:
                                st.sidebar.error("Update failed. Check logs or details above.")
                        else:
                            st.sidebar.warning("Please select an item and a valid shipped date.")
                except ValueError:
                    st.sidebar.error("Invalid item format selected. Please re-select.")
            elif unshipped_items_df.empty and not data_df.empty:
                 st.sidebar.info("All items are marked as shipped or data is unavailable for shipping status.")
        else:
            st.sidebar.info("No items available to mark as shipped, or all items are already shipped.")

    else:
        st.sidebar.warning("RMA, S/N, or Shipped columns not found in data. Cannot update shipping status.")


    # --- Display Filtered Data ---
    st.subheader("Filtered Data View")
    st.markdown(f"Displaying **{len(filtered_df)}** records out of **{len(data_df) if not data_df.empty else 0}** total records.")
    
    if not filtered_df.empty:
        # Display all columns as per EXPECTED_COLUMN_ORDER that are present in filtered_df
        display_cols = [col for col in EXPECTED_COLUMN_ORDER if col in filtered_df.columns]
        st.dataframe(filtered_df[display_cols].astype(str), use_container_width=True)
    else:
        st.warning("No data matches the current filter criteria or no data loaded.")

    if not filtered_df.empty:
        st.sidebar.markdown("---")
        st.sidebar.subheader("Download Data")
        # Ensure CSV download uses the filtered and correctly ordered data
        display_cols_download = [col for col in EXPECTED_COLUMN_ORDER if col in filtered_df.columns]
        csv = filtered_df[display_cols_download].to_csv(index=False).encode('utf-8')
        st.sidebar.download_button(
            label="Download Filtered Data as CSV",
            data=csv,
            file_name='filtered_service_data.csv',
            mime='text/csv',
        )
else:
    st.info("No data to display. Please ensure the Google Sheet is accessible, contains data with headers for all expected columns, and 'Credentials.json' is correctly set up.")

st.markdown("---")
st.markdown("Built with ❤️ using [Streamlit](https://streamlit.io) and Google Sheets")
