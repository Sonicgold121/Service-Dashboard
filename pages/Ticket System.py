# pages/Ticket_System.py

import streamlit as st
import pandas as pd
from logic import send_ticket_reply_and_log, update_ticket_status
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import urllib.parse

# --- Page Config ---
st.set_page_config(page_title="Ticketing System", layout="wide")
st.title("🎫 Customer Ticket System")

# --- Business Central Constants ---
BC_BASE_URL = "https://businesscentral.dynamics.com/7bcfb5b0-27a1-4e18-99d8-ca66570addd8/Production"
BC_COMPANY = "PROD"
BC_PAGE_ID = "70001"
BC_RMA_FIELD_NAME = "No."
BC_LINK_COL_NAME = "Business Central Link" # This will be the name of our new column

# --- Google Sheets Connection Functions (no changes here) ---
@st.cache_resource(ttl=300)
def connect_and_get_sheet():
    # ... (same function as before)
    try:
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
        client = gspread.authorize(creds)
        sheet = client.open("Estimate form").worksheet("Tickets")
        return sheet
    except Exception as e:
        st.error(f"Could not connect to Google Sheets: {e}")
        return None

@st.cache_data(ttl=60)
def load_tickets(_sheet):
    # ... (same function as before)
    if _sheet is None: return pd.DataFrame()
    records = _sheet.get_all_records()
    return pd.DataFrame(records)

# --- Main App ---
sheet = connect_and_get_sheet()

if sheet:
    df_tickets = load_tickets(sheet)

    if df_tickets.empty:
        st.info("No tickets found.")
    else:
        # --- NEW: DYNAMICALLY CREATE THE LINK COLUMN ---
        if 'RMA' in df_tickets.columns:
            df_tickets[BC_LINK_COL_NAME] = df_tickets['RMA'].apply(
                lambda rma: f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(str(rma))}%27"
                if pd.notna(rma) and str(rma).strip() != "" else None
            )

        # ... (the rest of the script is the same, but we refer to BC_LINK_COL_NAME) ...
        st.sidebar.header("Filter Tickets")
        # ... (filter logic is the same) ...

        st.dataframe(
            df_filtered, # Make sure to use the filtered dataframe
            column_order=("Ticket ID", "Status", "RMA", BC_LINK_COL_NAME, "Customer Email", "Subject"),
            column_config={
                BC_LINK_COL_NAME: st.column_config.LinkColumn(
                    "View in BC",
                    display_text="Open Link"
                )
            },
            width='stretch',
            hide_index=True
        )

        # ... (All the rest of your page logic for replying and closing tickets is the same) ...
