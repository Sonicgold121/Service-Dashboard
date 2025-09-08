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
        
        # Get all values from the sheet (more robust for empty columns with headers)
        all_values = worksheet.get_all_values()
        
        if not all_values:
            st.warning(f"No data (not even headers) found in Google Sheet '{sheet_name}', worksheet index {worksheet_index}.")
            return pd.DataFrame()
            
        headers = all_values[0]
        data_rows = all_values[1:]
        
        df = pd.DataFrame(data_rows, columns=headers)
        
        if df.empty and headers: # Handles case where there are headers but no data rows
             st.info(f"Google Sheet '{sheet_name}', worksheet index {worksheet_index} has headers but no data rows.")
             # We can proceed with an empty DataFrame that has the correct columns for type conversion
             # This ensures downstream code doesn't break if it expects these columns.

        # --- Data Cleaning and Type Conversion ---
        # Define columns that should be treated primarily as strings and where "N/A" is appropriate for empty/missing
        # This includes descriptive fields and status fields.
        string_cols_for_na_fill = ['RMA', 'S/N', 'Part Number', 'SPC Code', 
                                   'Description', 'Fault Comments', 'Resolution Comments', 
                                   'Sender', 'Estimate Complete', 'Estimate Approved', 'QA Approved']
        for col in string_cols_for_na_fill:
            if col in df.columns:
                df[col] = df[col].astype(str) # Ensure column is string type
                # Replace common placeholders for empty or null-like strings with "N/A"
                df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None], 'N/A') 
            elif col in headers: # If column was in header but not in df (e.g. no data rows)
                df[col] = "N/A" # Create the column and fill with N/A


        # Convert date columns to datetime objects, coercing errors
        # Values that cannot be parsed (including "N/A" if it was an empty string and got converted above)
        # will become NaT (Not a Time).
        date_cols = ['Estimate Complete Time', 'Estimate Approved Time', 'QA Approved Time']
        for col in date_cols:
            if col in df.columns:
                # Replace "N/A" (from previous step if it was originally empty) with None before to_datetime
                # so that it correctly becomes NaT.
                df[col] = df[col].replace('N/A', None) 
                df[col] = pd.to_datetime(df[col], errors='coerce')
            elif col in headers: # If column was in header but not in df
                 df[col] = pd.NaT # Create the column as datetime and fill with NaT


        # Ensure all expected columns are present, even if they were entirely empty and not in string_cols_for_na_fill or date_cols
        # This is a fallback, ideally all relevant columns are handled above.
        for header_col in headers:
            if header_col not in df.columns:
                df[header_col] = "N/A" # Default fill for any other missed columns

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
        st.error(f"Details: {type(e).__name__} - {str(e)}") # More detailed error
        return pd.DataFrame()

def display_kpis(df):
    """Displays Key Performance Indicators."""
    if df.empty:
        return

    total_records = len(df)
    
    # Ensure status columns exist before trying to count and that comparison is with "Yes"
    estimate_complete_count = df[df['Estimate Complete'].astype(str) == 'Yes'].shape[0] if 'Estimate Complete' in df.columns else 0
    estimate_approved_count = df[df['Estimate Approved'].astype(str) == 'Yes'].shape[0] if 'Estimate Approved' in df.columns else 0
    qa_approved_count = df[df['QA Approved'].astype(str) == 'Yes'].shape[0] if 'QA Approved' in df.columns else 0


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
    # Clear Streamlit's internal cache for data loading functions
    # This ensures load_data_from_google_sheet is re-run.
    # Use st.cache_data.clear() if load_data_from_google_sheet is decorated with @st.cache_data
    # If not decorated, st.rerun() is enough to trigger a reload.
    # Forcing a clear if you suspect caching issues with gspread or underlying libraries:
    try:
        st.cache_data.clear()
    except: # Broad except as the function might not exist if streamlit version is old
        pass
    st.rerun()


# For simplicity, using hardcoded values from your script
data_df = load_data_from_google_sheet(sheet_name="Estimate form", worksheet_index=1, creds_file="Credentials.json")


if not data_df.empty:
    st.subheader("📊 Key Metrics")
    display_kpis(data_df.copy()) # Pass a copy to avoid modifying the original df
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
            # Get unique values, ensure 'All' is an option. Values should be strings, including "N/A".
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
            # Ensure the column is in datetime format for min/max and filtering
            temp_col_dt = filtered_df[col_name] # Should be datetime64[ns] with NaT for errors

            if temp_col_dt.notna().any(): # Check if there are any valid dates (non-NaT)
                min_date_val = temp_col_dt.min() # Will be NaT if all are NaT
                max_date_val = temp_col_dt.max() # Will be NaT if all are NaT

                if pd.isna(min_date_val) or pd.isna(max_date_val): # Check if min or max is NaT
                    st.sidebar.warning(f"Not enough valid date data in '{display_name}' to create a range filter.")
                    continue
                
                # Convert pd.Timestamp to datetime.date for st.date_input widget
                min_date_dt = min_date_val.date() 
                max_date_dt = max_date_val.date()
                
                try:
                    date_range = st.sidebar.date_input(
                        f"Filter by {display_name}",
                        value=(min_date_dt, max_date_dt), # Use date objects
                        min_value=min_date_dt,
                        max_value=max_date_dt,
                        key=f"date_range_{col_name}"
                    )
                    if date_range and len(date_range) == 2:
                        start_date, end_date = date_range
                        # Convert selected dates from widget (which are date objects) to datetime for comparison
                        start_datetime = pd.to_datetime(start_date)
                        end_datetime = pd.to_datetime(end_date).replace(hour=23, minute=59, second=59)
                        
                        # The column in filtered_df is already datetime64[ns] from load_data
                        # Apply filter, NaT values will not satisfy the condition.
                        filtered_df = filtered_df[
                            (filtered_df[col_name] >= start_datetime) & 
                            (filtered_df[col_name] <= end_datetime) &
                            (filtered_df[col_name].notna()) # Explicitly include only non-NaT dates
                        ]
                except Exception as e:
                    st.sidebar.error(f"Error with date filter for {display_name}: {e}")
            else:
                st.sidebar.warning(f"No valid date data in '{display_name}' for filtering.")


    st.subheader("Filtered Data View")
    st.markdown(f"Displaying **{len(filtered_df)}** records out of **{len(data_df) if not data_df.empty else 0}** total records.")
    
    if not filtered_df.empty:
        # Convert all columns to string for display to avoid Streamlit's issues with mixed types or Arrow conversion
        # This also converts NaT to "NaT" and NaN to "nan"
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
    st.info("No data to display. Please ensure the Google Sheet is accessible, contains data with headers for all expected columns, and 'Credentials.json' is correctly set up.")

st.markdown("---")
st.markdown("Built with ❤️ using [Streamlit](https://streamlit.io)")
