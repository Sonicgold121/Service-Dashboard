# In logic.py

def send_ticket_reply_and_log(sheet, ticket_id, customer_email, original_subject, reply_body, team_member_name):
    """
    Sends an email reply, logs it, and searches the reply for an RMA number
    to update the ticket record.
    """
    try:
        # --- Part 1: Send the email via Resend (No changes here) ---
        resend.api_key = st.secrets["resend"]["api_key"]
        
        full_reply_html = f"""
        <p>{reply_body.replace('\\n', '<br>')}</p>
        <br><p>--- Original Message ---</p><blockquote>{original_subject}</blockquote>
        """

        params = {
            "from": f"{team_member_name} <onboarding@resend.dev>",
            "to": [customer_email],
            "subject": f"Re: {original_subject}",
            "html": full_reply_html,
        }
        
        email = resend.Emails.send(params)
        
        # --- Part 2: Log the reply and find the Ticket Row (No changes here) ---
        cell = sheet.find(ticket_id)
        if not cell:
            return False, f"Could not find ticket {ticket_id} in the sheet to log the reply."
        
        row_index = cell.row
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        note = f"--- Reply Sent by {team_member_name} at {timestamp} ---\\n{reply_body}\\n\\n"
        
        notes_col_index = sheet.find("Notes").col
        existing_notes = sheet.cell(row_index, notes_col_index).value or ""
        updated_notes = note + existing_notes
        
        sheet.update_cell(row_index, notes_col_index, updated_notes)
        sheet.update_cell(row_index, sheet.find("Status").col, "In Progress")

        # --- PART 3: NEW - Find RMA and update the sheet ---
        # This pattern now looks for the specific "RMA" + numbers format.
        rma_match = re.search(r'(RMA\d+)', reply_body, re.IGNORECASE)
        
        if rma_match:
            rma_number = rma_match.group(1) # Get the captured number (e.g., "RMA01234")
            
            # Find the columns to update
            rma_col_index = sheet.find("RMA").col
            bc_link_col_index = sheet.find("Business Central Link").col
            
            # Use the constants from your dashboard to create the link
            bc_base_url = "https://businesscentral.dynamics.com/7bcfb5b0-27a1-4e18-99d8-ca66570addd8/Production"
            bc_company = "PROD"
            bc_page_id = "70001"
            bc_rma_field_name = "No."
            
            bc_link = f"{bc_base_url}?company={bc_company}&page={bc_page_id}&filter='{urllib.parse.quote_plus(bc_rma_field_name)}'%20IS%20%27{urllib.parse.quote_plus(rma_number)}%27"
            
            # Update the cells in the Google Sheet
            sheet.update_cell(row_index, rma_col_index, rma_number)
            sheet.update_cell(row_index, bc_link_col_index, bc_link)

        return True, "Successfully sent reply and updated ticket log."

    except Exception as e:
        return False, f"An error occurred: {e}"
else:
    st.error("Failed to connect to the Google Sheet.")
    st.warning("The page cannot display tickets without a connection to the 'Tickets' worksheet. Please check the troubleshooting steps.")

