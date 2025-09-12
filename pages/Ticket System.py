import streamlit as st
import pandas as pd
from logic import send_ticket_reply_and_log # We will create this function next
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- Page Config ---
st.set_page_config(page_title="Ticketing System", layout="wide")
st.title("🎫 Customer Ticket System")

# --- Google Sheets Connection (Cached) ---
@st.cache_resource(ttl=300)
def connect_and_get_sheet():
    try:
        # 'scopes' is defined here
        scopes = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                  "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scopes) # and used here
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Failed to connect to Google Sheets: {e}")
        return Nonee

@st.cache_data(ttl=60)
def load_tickets(_sheet):
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
        st.sidebar.header("Filter Tickets")
        status_filter = st.sidebar.selectbox("Filter by Status", options=["All"] + df_tickets["Status"].unique().tolist())

        if status_filter != "All":
            df_filtered = df_tickets[df_tickets["Status"] == status_filter]
        else:
            df_filtered = df_tickets

        st.dataframe(df_filtered, use_container_width=True)
        st.markdown("---")

        st.header("Reply to a Ticket")
        
        # Create a list of tickets for the selectbox
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
            
            with st.form(key="reply_form"):
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
                                team_member_name="Service Team" # Or get from logged-in user later
                            )
                            if success:
                                st.success(message)
                                # Clear cache to show updated notes
                                load_tickets.clear()
                            else:

                                st.error(message)

