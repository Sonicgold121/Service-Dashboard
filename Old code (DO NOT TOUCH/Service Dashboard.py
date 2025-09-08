import streamlit as st
import pandas as pd
from datetime import datetime
import gspread # Added for Google Sheets
from oauth2client.service_account import ServiceAccountCredentials # Added for Google Sheets

# --- Page Configuration ---
st.set_page_config(
    page_title="Email Data Dashboard",
    page_icon="📧",
    layout="wide",
)

# --- Helper Functions ---
# @st.cache_data(ttl=600) # Optional: Cache data for 10 minutes
def load_data_from_google_sheet(
    sheet_name="Estimate form", 
    worksheet_index=1, 
    creds_file="Credentials.json"
):
    """Loads data from the specified Google Sheet."""
    try:
        # Google Sheets API setup
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
        client = gspread.authorize(creds)

        # Open the Google Sheet
        spreadsheet = client.open(sheet_name)
        
        # Access the specified worksheet (e.g., the second sheet by index 1)
        worksheet = spreadsheet.get_worksheet(worksheet_index)
        
        # Get all records from the sheet
        records = worksheet.get_all_records() # This returns a list of dictionaries
        
        if not records:
            st.warning(f"No data found in Google Sheet '{sheet_name}', worksheet index {worksheet_index}.")
            return pd.DataFrame()
            
        df = pd.DataFrame(records)

        # --- Data Cleaning and Type Conversion (similar to original load_data) ---
        # Convert relevant columns to string
        for col in ['RMA', 'S/N', 'Part Number', 'SPC Code']:
            if col in df.columns:
                # gspread might return empty strings as '', not NaN, so handle accordingly
                df[col] = df[col].astype(str).replace({'nan': '', 'None': ''}).fillna('')


        # Convert date columns to datetime objects, coercing errors
        # Google Sheets might store dates as strings; ensure they are parsed correctly.
        date_cols = ['Estimate Complete Time', 'Estimate Approved Time', 'QA Approved Time']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Fill NaN/None values for display purposes (after all type conversions)
        # df.fillna("N/A", inplace=True) # Be careful with this if you need to distinguish N/A from empty
        # Replace empty strings with "N/A" for display if desired, or handle in display logic
        df.replace('', "N/A", inplace=True) # Replace empty strings from gspread with N/A
        df.fillna("N/A", inplace=True) # Catch any remaining NaNs

        return df

    except FileNotFoundError:
        st.error(f"Error: Credentials file '{creds_file}' not found. Please ensure it's in the correct path.")
        return pd.DataFrame()
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"Error: Google Sheet '{sheet_name}' not found. Please check the name and permissions.")
        return pd.DataFrame()
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Error: Worksheet with index {worksheet_index} not found in '{sheet_name}'.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"An error occurred while loading data from Google Sheets: {e}")
        return pd.DataFrame()

def display_kpis(df):
    """Displays Key Performance Indicators."""
    if df.empty:
        return

    total_records = len(df)
    
    estimate_complete_count = df[df['Estimate Complete'] == 'Yes'].shape[0] if 'Estimate Complete' in df.columns else 0
    estimate_approved_count = df[df['Estimate Approved'] == 'Yes'].shape[0] if 'Estimate Approved' in df.columns else 0
    qa_approved_count = df[df['QA Approved'] == 'Yes'].shape[0] if 'QA Approved' in df.columns else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Records", total_records)
    col2.metric("Estimates Complete", estimate_complete_count)
    col3.metric("Estimates Approved", estimate_approved_count)
    col4.metric("QA Approved", qa_approved_count)

# --- Main Application ---
st.title("📧 Email Data Monitoring Dashboard (from Google Sheets)")
st.markdown("This dashboard displays data extracted from emails, updated via Google Sheets.")

# --- Load Data ---
if st.button("🔄 Refresh Data"):
    st.cache_data.clear() # Clear cache if using st.cache_data for load_data_from_google_sheet
    st.rerun()

# --- User Inputs for Google Sheet (Optional, or hardcode) ---
# You can hardcode these or allow user input
# default_sheet_name = "Estimate form"
# default_worksheet_index = 1 # 0 for first sheet, 1 for second, etc.
# default_creds_file = "Credentials.json"

# For simplicity, using hardcoded values from your script
data_df = load_data_from_google_sheet(sheet_name="Estimate form", worksheet_index=1, creds_file="Credentials.json")


if not data_df.empty:
    st.subheader("📊 Key Metrics")
    display_kpis(data_df.copy())
    st.markdown("---")

    st.sidebar.header("🔍 Filter Options")
    filtered_df = data_df.copy()

    # Text search
    for col_name, search_label in [('RMA', "RMA"), ('S/N', "S/N"), 
                                   ('Part Number', "Part Number"), ('SPC Code', "SPC Code")]:
        if col_name in filtered_df.columns:
            search_term = st.sidebar.text_input(f"Search by {search_label}", key=f"search_{col_name}")
            if search_term:
                # Ensure comparison is with string version of column
                filtered_df = filtered_df[filtered_df[col_name].astype(str).str.contains(search_term, case=False, na=False)]
    
    # Dropdown filters for status columns
    status_columns = {
        'Estimate Complete': 'Estimate Complete',
        'Estimate Approved': 'Estimate Approved',
        'QA Approved': 'QA Approved'
    }
    for display_name, col_name in status_columns.items():
        if col_name in filtered_df.columns:
            # Handle cases where a status might be "N/A" or other placeholders from fillna
            unique_values = ['All'] + sorted(filtered_df[col_name].astype(str).unique().tolist())
            selected_status = st.sidebar.selectbox(f"Filter by {display_name}", unique_values, key=f"select_{col_name}")
            if selected_status != "All":
                filtered_df = filtered_df[filtered_df[col_name].astype(str) == selected_status]

    # Date range filters
    st.sidebar.markdown("---")
    st.sidebar.subheader("Date Range Filters")
    date_filter_columns = {
        'Estimate Complete Time': 'Estimate Complete Time',
        'Estimate Approved Time': 'Estimate Approved Time',
        'QA Approved Time': 'QA Approved Time'
    }

    for display_name, col_name in date_filter_columns.items():
        if col_name in filtered_df.columns:
            # Make sure column is datetime before attempting min/max
            temp_col_dt = pd.to_datetime(filtered_df[col_name], errors='coerce')
            
            if not temp_col_dt.empty and temp_col_dt.notna().any(): # Check if there are any valid dates
                min_date_val = temp_col_dt.min()
                max_date_val = temp_col_dt.max()

                if pd.isna(min_date_val) or pd.isna(max_date_val):
                    st.sidebar.warning(f"Not enough valid date data in '{display_name}' to create a range filter.")
                    continue
                
                min_date_dt = min_date_val.date() if isinstance(min_date_val, (pd.Timestamp, datetime)) else datetime.min.date()
                max_date_dt = max_date_val.date() if isinstance(max_date_val, (pd.Timestamp, datetime)) else datetime.max.date()
                
                try:
                    date_range = st.sidebar.date_input(
                        f"Filter by {display_name}",
                        value=(min_date_dt, max_date_dt),
                        min_value=min_date_dt,
                        max_value=max_date_dt,
                        key=f"date_range_{col_name}"
                    )
                    if date_range and len(date_range) == 2:
                        start_date, end_date = date_range
                        start_datetime = pd.to_datetime(start_date)
                        end_datetime = pd.to_datetime(end_date).replace(hour=23, minute=59, second=59)
                        
                        # Ensure the column being filtered is also datetime
                        filtered_df[col_name] = pd.to_datetime(filtered_df[col_name], errors='coerce')
                        # Drop rows where date conversion failed if necessary for filtering
                        # filtered_df.dropna(subset=[col_name], inplace=True) 
                        
                        # Apply filter, handling NaT (Not a Time) values by not including them unless explicitly handled
                        is_valid_date = filtered_df[col_name].notna()
                        filtered_df = filtered_df[
                            is_valid_date &
                            (filtered_df[col_name] >= start_datetime) & 
                            (filtered_df[col_name] <= end_datetime)
                        ]
                except Exception as e:
                    st.sidebar.error(f"Error with date filter for {display_name}: {e}")
            else:
                st.sidebar.warning(f"No valid date data in '{display_name}' for filtering.")


    st.subheader("Filtered Data View")
    st.markdown(f"Displaying **{len(filtered_df)}** records out of **{len(data_df) if not data_df.empty else 0}** total records.")
    
    if not filtered_df.empty:
        st.dataframe(filtered_df.astype(str), use_container_width=True)
    else:
        st.warning("No data matches the current filter criteria or no data loaded.")

    if not filtered_df.empty:
        st.sidebar.markdown("---")
        st.sidebar.subheader("Download Data")
        csv = filtered_df.to_csv(index=False).encode('utf-8')
        st.sidebar.download_button(
            label="Download Filtered Data as CSV",
            data=csv,
            file_name='filtered_google_sheet_data.csv',
            mime='text/csv',
        )
else:
    st.info("No data to display. Please ensure the Google Sheet is accessible, contains data, and 'Credentials.json' is correctly set up.")

st.markdown("---")
st.markdown("Built with ❤️ using [Streamlit](https://streamlit.io)")