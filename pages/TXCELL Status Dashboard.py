import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta 
import gspread 
from oauth2client.service_account import ServiceAccountCredentials 
import urllib.parse 
from io import BytesIO

# --- Page Configuration ---
st.set_page_config(
    page_title="TXCELL Status Dashboard", 
    page_icon="🔬", 
    layout="wide",
)

# --- Constants for Google Sheets & Business Central ---
GSHEET_NAME = "Estimate form"
WORKSHEET_INDEX = 1 # Main data sheet
CREDS_FILE = "Credentials.json" 

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

BC_BASE_URL = "https://businesscentral.dynamics.com/7bcfb5b0-27a1-4e18-99d8-ca66570addd8/Production"
BC_COMPANY = "PROD"
BC_PAGE_ID = "9318"  # <-- Updated Page ID for Service Orders
BC_RMA_FIELD_NAME = "RSMUS SDM ServReq No." # <-- Updated Field Name
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
            return pd.DataFrame(columns=EXPECTED_COLUMN_ORDER) 
            
        headers_from_sheet = all_values[0]
        data_rows = all_values[1:]
        
        temp_df = pd.DataFrame(data_rows, columns=headers_from_sheet)
        df = pd.DataFrame(columns=EXPECTED_COLUMN_ORDER)

        for col in EXPECTED_COLUMN_ORDER:
            if col in temp_df.columns:
                df[col] = temp_df[col] 
            else: 
                 if "Time" in col: df[col] = pd.NaT
                 elif col in ALL_STATUS_COLUMNS: df[col] = "No"
                 else: df[col] = "N/A" 
        
        df = df[EXPECTED_COLUMN_ORDER] 

        string_cols_to_process = ['RMA', 'S/N', 'Part Number', 'SPC Code', 
                                   'Description', 'Fault Comments', 'Resolution Comments', 
                                   'Sender'] + ALL_STATUS_COLUMNS
        for col in string_cols_to_process:
            if col in df.columns:
                df[col] = df[col].astype(str) 
                if col in ALL_STATUS_COLUMNS:
                     df[col] = df[col].replace(['', 'nan', 'None', 'NaN', 'NONE', None, 'NaT'], 'No') 
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

def create_excel_report(awaiting_qa_df, needs_approval_df):
    """Creates an Excel file with sheets for items awaiting QA and needing approval."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
        
        needs_approval_df.to_excel(writer, sheet_name='Needs Approval', index=False)
        worksheet_not_approved = writer.sheets['Needs Approval']
        for col_num, value in enumerate(needs_approval_df.columns.values):
            worksheet_not_approved.write(0, col_num, value, header_format)
            if not needs_approval_df.empty:
                max_len = max(needs_approval_df[value].astype(str).map(len).max(), len(value)) + 2
                worksheet_not_approved.set_column(col_num, col_num, max_len)
            
        awaiting_qa_df.to_excel(writer, sheet_name='Approved (Awaiting QA)', index=False)
        worksheet_approved = writer.sheets['Approved (Awaiting QA)']
        for col_num, value in enumerate(awaiting_qa_df.columns.values):
            worksheet_approved.write(0, col_num, value, header_format)
            if not awaiting_qa_df.empty:
                max_len = max(awaiting_qa_df[value].astype(str).map(len).max(), len(value)) + 2
                worksheet_approved.set_column(col_num, col_num, max_len)

    return output.getvalue()

# --- Main Application ---
st.title("TXCELL Approval Status")
st.markdown("This page tracks the approval and QA status for all service items with the part number `CUST-TXCELL`.")

if st.button("🔄 Refresh Data"):
    st.cache_data.clear()
    st.rerun()

data_df = load_data_from_google_sheet()

if data_df.empty:
    st.warning("Could not load data from the Google Sheet. Please check the connection and sheet contents.")
else:
    filtered_df = data_df[data_df['Part Number'] == 'CUST-TXCELL'].copy()

    if filtered_df.empty:
        st.info("✅ No service records found with the part number 'CUST-TXCELL'.")
    else:
        filtered_df[BC_LINK_COL_NAME] = filtered_df.apply(
            lambda row: f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(str(row['RMA']))}%27"
            if pd.notna(row['RMA']) and str(row['RMA']).strip() != 'N/A' and str(row['RMA']).strip() != "" else None, axis=1
        )
        
        needs_approval_df = filtered_df[filtered_df['Estimate Approved'].str.lower() == 'no']
        
        awaiting_qa_df = filtered_df[
            (filtered_df['Estimate Approved'].str.lower() == 'yes') & 
            (filtered_df['QA Approved'].str.lower() == 'no')
        ].copy()

        if not awaiting_qa_df.empty:
            today_dt = pd.to_datetime(date.today()).normalize()
            awaiting_qa_df['Days Since Approval'] = (today_dt - awaiting_qa_df['Estimate Approved Time'].dt.normalize()).dt.days
            awaiting_qa_df = awaiting_qa_df.sort_values(by='Estimate Approved Time', ascending=True)

        st.markdown("---")
        kpi1, kpi2, btn_col = st.columns([1, 1, 2])
        kpi1.metric(label="Waiting for Approval", value=len(needs_approval_df))
        kpi2.metric(label="Approved & Awaiting QA", value=len(awaiting_qa_df))
        
        report_bytes = create_excel_report(awaiting_qa_df, needs_approval_df)
        with btn_col:
            st.write("") 
            st.download_button(
                label="📄 Download TXCELL Report",
                data=report_bytes,
                file_name=f"CUST-TXCELL_Action_Report_{date.today().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        st.markdown("---")

        st.subheader("Estimates Not Approved")
        if not needs_approval_df.empty:
            display_cols_not_approved = ['RMA', BC_LINK_COL_NAME, 'S/N', 'Description', 'Fault Comments', 'Sender', 'Estimate Complete Time', 'Estimate Sent Time']
            st.dataframe(needs_approval_df[display_cols_not_approved], use_container_width=True, column_config={BC_LINK_COL_NAME: st.column_config.LinkColumn(label="Business Central", display_text="Open RMA")})
        else:
            st.success("✅ All CUST-TXCELL estimates have an approval status.")

        st.markdown("---")

        st.subheader("Estimate Approved")
        if not awaiting_qa_df.empty:
            display_cols_approved = ['RMA', BC_LINK_COL_NAME, 'S/N', 'Description', 'Sender', 'Estimate Approved Time', 'Days Since Approval']
            st.dataframe(awaiting_qa_df[display_cols_approved], use_container_width=True, column_config={BC_LINK_COL_NAME: st.column_config.LinkColumn(label="Business Central", display_text="Open RMA"), 'Estimate Approved Time': st.column_config.DateColumn("Approved Date", format="YYYY-MM-DD")})
        else:
            st.success("✅ No CUST-TXCELL items are currently awaiting QA.")