# pages2_Search_Estimate.py

import streamlit as st
import os
import glob
from datetime import date
import pandas as pd
import openpyxl # Import openpyxl to read specific cells

# Import the necessary logic functions
from logic import send_estimate_email, update_estimate_sent_details_in_gsheet

# --- Page Configuration ---
st.set_page_config(layout="wide")
st.title("🔎 Search & Resend Past Estimates")
st.write("Search for previously generated estimate files by RMA Number.")

# --- Define the directory where estimates are saved ---
ESTIMATE_DIRECTORY = "generated_estimates"
if not os.path.exists(ESTIMATE_DIRECTORY):
    os.makedirs(ESTIMATE_DIRECTORY)

# --- Search Bar ---
search_query = st.text_input("Enter RMA Number to search for files", key="search_query")
search_button = st.button("Search Files")

# --- Search Logic ---
if search_button:
    if not search_query:
        st.warning("Please enter an RMA Number to search.")
        if 'found_files' in st.session_state:
            del st.session_state['found_files']
    else:
        with st.spinner(f"Searching for files containing '{search_query}'..."):
            search_pattern = os.path.join(ESTIMATE_DIRECTORY, f"*{search_query}*.*")
            found_files = glob.glob(search_pattern)
            
            st.session_state['found_files'] = found_files
            st.session_state['last_search'] = search_query
            st.rerun()

# --- Display Results & Resend Section ---
if 'found_files' in st.session_state:
    found_files = st.session_state['found_files']
    last_search = st.session_state.get('last_search', '')
    
    st.write("---")
    if found_files:
        st.success(f"Found **{len(found_files)}** matching file(s) for '{last_search}'.")
        
        for file_path in found_files:
            file_name = os.path.basename(file_path)
            
            col1, col2 = st.columns([3, 1])
            with col1:
                st.markdown(f"📄 **{file_name}**")
            with col2:
                try:
                    with open(file_path, "rb") as fp:
                        st.download_button(
                            label="Download",
                            data=fp,
                            file_name=file_name,
                            mime='application/octet-stream',
                            key=f"download_{file_name}"
                        )
                except FileNotFoundError:
                    st.error("File not found.")

            # --- Section to display parts and price from Excel file ---
            if file_path.lower().endswith('.xlsx'):
                with st.expander("View Parts & Price Details"):
                    try:
                        # FIX: Use openpyxl to read specific cell ranges
                        workbook = openpyxl.load_workbook(file_path, data_only=True)
                        sheet = workbook.active
                        
                        parts_data = {
                            'Part Number': [],
                            'Description': [],
                            'Quantity': [],
                            'Price/Unit': [],
                            'Total Price': []
                        }
                        
                        # Loop from row 19 to 31
                        for row in range(19, 32):
                            part_num = sheet[f'A{row}'].value
                            # If Part Number is empty, assume it's the end of the list
                            if part_num is None or str(part_num).strip() == "":
                                continue
                                
                            parts_data['Part Number'].append(part_num)
                            parts_data['Price/Unit'].append(sheet[f'B{row}'].value)
                            parts_data['Quantity'].append(sheet[f'C{row}'].value)
                            parts_data['Description'].append(sheet[f'E{row}'].value)
                            parts_data['Total Price'].append(sheet[f'I{row}'].value)

                        if parts_data['Part Number']: # Check if any parts were found
                            display_df = pd.DataFrame(parts_data)
                            st.dataframe(display_df)
                            
                            # Calculate and display the total cost
                            total_cost = pd.to_numeric(display_df['Total Price'], errors='coerce').sum()
                            st.metric("Total Estimated Cost", f"${total_cost:,.2f}")
                        else:
                            st.info("No parts information found in the specified range (A19:I31) of this file.")

                    except Exception as e:
                        st.error(f"Could not read details from this Excel file. Error: {e}")


            # Only show the resend option for PDF files
            if file_path.lower().endswith('.pdf'):
                with st.expander("Resend this Estimate via Email"):
                    try:
                        rma_from_filename = file_name.split('_')[-1].split('.')[0]
                        sn_from_gsheet = "N/A" 
                    except IndexError:
                        rma_from_filename = last_search

                    recipient = st.text_input("Recipient Email Address", key=f"email_{file_name}")
                    if st.button("📧 Send Email", key=f"send_{file_name}"):
                        if recipient and "@" in recipient:
                            with st.spinner("Sending email..."):
                                email_success, _ = send_estimate_email(recipient, rma_from_filename, file_path)
                                if email_success:
                                    st.success(f"Email successfully sent to {recipient}!")
                                    update_estimate_sent_details_in_gsheet(rma_from_filename, sn_from_gsheet, recipient, date.today())
                                else:
                                    st.error("Failed to send email. Ensure Outlook is running.")
                        else:
                            st.warning("Please enter a valid email address.")
            st.markdown("---")

    elif last_search:
        st.error(f"No estimate files found containing '{last_search}' in the '{ESTIMATE_DIRECTORY}' folder.")
