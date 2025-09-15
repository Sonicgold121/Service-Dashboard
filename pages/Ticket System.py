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

# --- Business Central Constants for Link Generation ---
BC_BASE_URL = "https://businesscentral.dynamics.com/7bcfb5b0-27a1-4e18-99d8-ca66570addd8/Production"
BC_COMPANY = "PROD"
BC_PAGE_ID = "70001" 
BC_RMA_FIELD_NAME = "No."
BC_LINK_COL_NAME = "Business Central Link" 

# --- Google Sheets Connection Functions ---
@st.cache_resource(ttl=300)
def connect_and_get_sheet():
    """Connects to Google Sheets and returns the 'Tickets' worksheet object."""
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
    """Loads all ticket records from the worksheet into a DataFrame."""
    if _sheet is None:
        return pd.DataFrame()
    records = _sheet.get_all_records()
    return pd.DataFrame(records)

# --- Main App ---
sheet = connect_and_get_sheet()

if sheet:
    df_tickets = load_tickets(sheet)

    if df_tickets.empty:
        st.info("No tickets found.")
    else:
        # --- DYNAMICALLY CREATE THE LINK COLUMN ---
        if 'RMA' in df_tickets.columns:
            df_tickets[BC_LINK_COL_NAME] = df_tickets['RMA'].apply(
                lambda rma: f"{BC_BASE_URL}?company={BC_COMPANY}&page={BC_PAGE_ID}&filter='{urllib.parse.quote_plus(BC_RMA_FIELD_NAME)}'%20IS%20%27{urllib.parse.quote_plus(str(rma))}%27"
                if pd.notna(rma) and str(rma).strip() != "" else None
            )
        
        # --- Sidebar and Filtering ---
        st.sidebar.header("Filter Tickets")
        if "Status" in df_tickets.columns:
            status_filter = st.sidebar.selectbox("Filter by Status", options=["All"] + df_tickets["Status"].unique().tolist())
            if status_filter != "All":
                df_filtered = df_tickets[df_tickets["Status"] == status_filter]
            else:
                df_filtered = df_tickets
        else:
            st.error("The 'Tickets' sheet is missing a 'Status' column.")
            df_filtered = df_tickets

        # --- Display the Dataframe ---
        st.dataframe(
            df_filtered,
            column_order=("Ticket ID", "Status", "RMA","Serial Number", BC_LINK_COL_NAME, "Customer Email", "Subject"),
            column_config={
                BC_LINK_COL_NAME: st.column_config.LinkColumn(
                    "View in BC",
                    display_text="Open Link"
                )
            },
            width='stretch',
            hide_index=True
        )
        st.markdown("---")

        # --- Reply Section ---
        st.header("Reply to a Ticket")
        
        ticket_options = [f"{row['Ticket ID']}: {row['Subject']}" for index, row in df_filtered.iterrows()]
        selected_ticket_str = st.selectbox("Select a ticket to reply to", options=[""] + ticket_options)

        if selected_ticket_str:
            ticket_id_to_reply = selected_ticket_str.split(":")[0]
            ticket_data = df_tickets[df_tickets["Ticket ID"] == ticket_id_to_reply].iloc[0]

            st.subheader(f"Replying to Ticket: {ticket_data['Ticket ID']}")
            st.write(f"**From:** {ticket_data['Customer Email']}")
            st.write(f"**Subject:** {ticket_data['Subject']}")
            with st.expander("Original Message"):
                st.write(ticket_data['Body'])
            
            st.markdown("---")
            col1, col2, col3 = st.columns(3)
            with col1:
                if st.button("Mark as In Progress", use_container_width=True):
                    success, message = update_ticket_status(sheet, ticket_data['Ticket ID'], "In Progress")
                    if success:
                        st.success(message)
                        load_tickets.clear()
                    else:
                        st.error(message)
            with col2:
                if st.button("Close This Ticket", type="primary", use_container_width=True):
                    success, message = update_ticket_status(sheet, ticket_data['Ticket ID'], "Closed")
                    if success:
                        st.success(message)
                        load_tickets.clear()
                    else:
                        st.error(message)

            with st.form(key="reply_form"):
                team_member = st.selectbox(
                    "Select your name (for email signature and logs)",
                    options=st.secrets.get("users", {}).get("team_members", ["Default User"])
                )
                reply_text = st.text_area("Your Reply:", height=200)
                submitted = st.form_submit_button("Send Reply")

                if submitted:
                    if not reply_text:
                        st.warning("Reply cannot be empty.")
                    else:
                        with st.spinner("Sending and logging reply..."):
                            success, message = send_ticket_reply_and_log(
                                sheet=sheet,
                                ticket_id=ticket_data['Ticket ID'],
                                customer_email=ticket_data['Customer Email'],
                                original_subject=ticket_data['Subject'],
                                reply_body=reply_text,
                                team_member_name=team_member
                            )
                            if success:
                                st.success(message)
                                load_tickets.clear()
                            else:
                                st.error(message)
# This 'else' block prevents a blank page if the connection fails
else:
    st.error("Failed to connect to the Google Sheet.")
    st.warning("The page cannot display tickets without a connection to the 'Tickets' worksheet. Please check sharing permissions and sheet name.")

