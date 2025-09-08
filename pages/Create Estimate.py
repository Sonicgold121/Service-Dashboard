# pages/1_📝_Create_Estimate.py

import streamlit as st
import pandas as pd
import os
from datetime import date
from io import BytesIO
import streamlit_shadcn_ui as sui
from logic import (
    generate_estimate_files,
    add_or_update_estimate_in_gsheet,
    send_estimate_email,
    update_estimate_sent_details_in_gsheet,
    load_price_library,
    update_price_library_and_usage_count,
    load_price_library_df,
    save_price_library_df,
    load_estimate_for_revision,
    get_revision_rma,
    load_customer_list,
    load_zone_ranges,
    load_shipping_prices,
    SOURCE_PARTS_ARCHIVE_DIR

)

# --- Page Configuration ---
st.set_page_config(layout="wide")
st.title("📝 Create a New Customer Estimate")

# ===================================================================
# HOW-TO GUIDE & LIVE DEMO BUTTON
# ===================================================================




# --- Define a directory to save generated files ---
SAVE_DIRECTORY = "generated_estimates"
if not os.path.exists(SAVE_DIRECTORY):
    os.makedirs(SAVE_DIRECTORY)

form_keys = [
    'rma', 'cust_name', 'cust_num', 'serial', 'contact',
    'description', 'evaluation', 'shipping_zip', 'shipping_item_type'
]
for key in form_keys:
    if key not in st.session_state:
        st.session_state[key] = ""

# --- Clear Form Callback Function ---
def clear_form_state():
    '''Clears all input fields and uploaded files from the session state.'''
    keys_to_clear = [
        'rma', 'cust_name', 'cust_num', 'serial', 'contact',
        'description', 'evaluation', 'parts_df', 'file_paths',
        'cc_form_path', 'uploader_key', 'file_uploader', 'rma_for_history',
        'parts_for_library',
        'shipping_zip',
        'shipping_item_type',
        'shipping_cost',
        'revision_data',
        'search_rma_input',
        'uploaded_file'
    ]
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]
    
    # This is the crucial line that forces the file uploader to reset
    st.session_state.uploader_key = st.session_state.get('uploader_key', 0) + 1

with st.expander("📖 How to Use This Page (Click to Expand)"):
    st.markdown("""
        This page helps you create, revise, and send customer estimates. Here’s the standard workflow:

        **➡️ Step 1: Fill in Customer & Shipping Details**
        -   Enter the **RMA No.** and other customer details at the top.
        -   **Auto-fill Customer Name:** If you have a customer saved in your 'Customer List' sheet, just type their **Customer Number** and press Enter. The customer's name should pop in automatically.
        -   **Calculate Shipping:** Enter the **Shipping ZIP Code** and select the **Item Type**. The shipping cost will be calculated and displayed.

        **➡️ Step 2: Upload the Parts Sheet**
        -   Use the **'Choose the parts Excel file'** uploader to select the Excel file containing the parts for the repair.
        -   The parts list will appear in an editable table below.
        -   **Automatic Price Corrections:**
            - If a part's price is missing in your file, it will be auto-filled from the Price Library.
            - If a part's price in your file is **higher** than the library price, it will be automatically corrected to the lower library price.
            - If your parts list includes a line item for **"BILLABLE FREIGHT"**, its price will be automatically set to your calculated shipping cost.

        **➡️ Step 3: Generate & Send**
        -   Once all details and prices are correct, click the **'Generate & Log Estimate'** button. This creates the official PDF and saves a record of the estimate.
        -   Finally, you can use the **'Send Email & Update Log'** section to email the PDF directly to the customer via your Outlook.

        ---
        #### **Other Features**
        * **Revising an Old Estimate:** Use the **'Revise an Existing Estimate'** section to search for and load a previous RMA to make corrections.
        * **Editing Prices:** The **'View and Edit Price Library'** section allows you to directly view and modify your saved part prices, descriptions, and usage counts.
    """)
    # --- NEW: Start Live Demo Button ---
    st.markdown("---")
    if st.button("🚀 Start Live Demo", type="primary"):
        clear_form_state() # Start with a clean slate
        st.session_state.demo_mode = True
        st.session_state.demo_step = 1 # Start at step 1
        st.rerun()

def populate_form_for_revision():
    '''Populates the form with data from a loaded estimate.'''
    data = st.session_state.get('revision_data')
    if data:
        # Using the exact keys from your debug output and wrapping every
        # line with str() to prevent data type errors.
        st.session_state['rma'] = get_revision_rma(str(data.get('RMA', '')))
        st.session_state['cust_name'] = str(data.get('Cust Name', ''))
        st.session_state['cust_num'] = str(data.get('Cust Num', ''))
        st.session_state['serial'] = str(data.get('S/N', ''))
        st.session_state['contact'] = str(data.get('Contact', ''))
        st.session_state['description'] = str(data.get('Customer Description of Problem', ''))
        
        # This is the corrected key including the invisible carriage return character
        st.session_state['evaluation'] = str(data.get('Technician Product Evaluate:\r', ''))
        
        # Handle based on record type
        if not data.get('is_legacy', True):
            st.session_state['parts_df'] = data.get('parts_df')
            st.session_state.uploader_key = st.session_state.get('uploader_key', 0) + 1
            if 'file_uploader' in st.session_state:
                del st.session_state['file_uploader']
        else:
            if 'parts_df' in st.session_state:
                del st.session_state['parts_df']

        del st.session_state['revision_data']

def autofill_customer_name():
    """Looks up the customer number and fills the name if a match is found."""
    customer_list = st.session_state.get('customer_list', {})
    cust_num = st.session_state.get('cust_num', '')
    if cust_num in customer_list:
        st.session_state['cust_name'] = customer_list[cust_num]

def run_live_demo():
    """Controls the step-by-step execution of the live demo."""
    step = st.session_state.get('demo_step', 0)
 
    
    # --- Demo UI Controls ---
    demo_container = st.container()
    with demo_container:
        st.info(f"**Live Demo: Step {step} of 4**", icon="🪄")
        col1, col2, _ = st.columns([1, 1, 3])
        
        if col1.button("Next Step", disabled=(step >= 4)):
            st.session_state.demo_step += 1
            st.rerun()

        if col2.button("End Demo & Clear Form", type="secondary"):
            st.session_state.demo_mode = False
            clear_form_state()
            st.rerun()

    # --- Demo Step Logic ---
    if step == 1 and st.session_state.rma == "":
        st.session_state.rma = "This is where the you write in the RMA. EX: 01234. "
        st.session_state.cust_num = " Type in the Customer number here. It will auto fill the customer name if the customer number is correct. EX: 10001" # IMPORTANT: Use a real Customer Number from your list for the best effect
        st.session_state.serial = "Type in the Unit serial here. Ex: IQI123456"
        st.session_state.contact = "Your name goes here"
        st.session_state.description = "Type in what was the issue that the customer sent"
        st.session_state.evaluation = "Type in what was the fault that the technican found"
        
        autofill_customer_name()
        st.toast("Step 1: Customer details filled!", icon="👨‍💼")
        st.rerun()
        
    elif step == 2 and st.session_state.shipping_zip == "":
        st.info(f"**Type in the customer Zipcode of where it going to ship**", icon="🪄")
        st.session_state.shipping_zip = "90210" # Example ZIP
        st.session_state.shipping_item_type = "Laser console"
        st.toast("Step 2: Shipping details filled!", icon="🚚")
        st.rerun()

    elif step == 3 and 'parts_df' not in st.session_state:
        try:
            mock_file_path = 'demo_parts.xlsx'
            demo_df = pd.read_excel(mock_file_path)
            demo_df['Amount Including Tax'] = demo_df['Amount Including Tax'].astype(float)
            
            if 'shipping_cost' in st.session_state and st.session_state.shipping_cost > 0:
                freight_mask = demo_df['No.'] == 'BILLABLE FREIGHT'
                if freight_mask.any():
                    demo_df.loc[freight_mask, 'Amount Including Tax'] = st.session_state.shipping_cost
            
            st.session_state.parts_df = demo_df
            st.toast("Step 3: Parts list 'uploaded' and populated!", icon="🧾")
            st.rerun()
        except FileNotFoundError:
            st.error("Could not find `demo_parts.xlsx`. Please create it to run the demo.")
            st.session_state.demo_mode = False
            
    elif step == 4:
        st.success("**Demo Complete!** The form is now filled out and ready to be generated.", icon="🎉")

if st.session_state.get('demo_mode', False):
    run_live_demo()

# --- Add a clear button at the top for easy access ---
st.button("Clear Form and Start Over", on_click=clear_form_state, type="secondary")
st.markdown("---")

with st.expander("📖 View and Edit Price Library"):
    st.info("Here you can directly view, add, delete, or edit entries in the Price Library. Click 'Save' to commit your changes to Google Sheets.")
    
    price_library_df = load_price_library_df()
    
    if not price_library_df.empty:
        # Use the data editor to allow changes.
        edited_df = st.data_editor(
            price_library_df,
            num_rows="dynamic", # Allows adding and deleting rows
            use_container_width=True,
            column_config={
                "No.": st.column_config.TextColumn("Part Number", required=True),
                "Description": st.column_config.TextColumn("Description"), # <-- ADD THIS
                "Amount Including Tax": st.column_config.NumberColumn(
                    "Unit Price (incl. Tax)",
                    format="$%.2f",
                    step=0.01,
                    required=True
                ),
                "Usage Count": st.column_config.NumberColumn(
                    "Usage Count",
                    step=1,
                    required=True
                )
            }
        )
        
        if st.button("Save Library Changes"):
            with st.spinner("Saving to Google Sheets..."):
                if save_price_library_df(edited_df):
                    st.success("✅ Success! The Price Library has been updated.")
                    # Clear caches to force a reload of the library data across the app
                    st.cache_data.clear()
                else:
                    st.error("❌ An error occurred while saving.")

    else:
        st.warning("Could not load the Price Library, or it is empty.")

with st.expander("🔄 Revise an Existing Estimate"):
    st.info("Enter an original RMA number to find and reload a previous estimate for correction.")
    search_rma = st.text_input("RMA to Find", key="search_rma_input")
    
    if st.button("Search for RMA"):
        if search_rma:
            with st.spinner(f"Searching for {search_rma}..."):
                # Clear previous search results before starting a new one
                if 'revision_data' in st.session_state:
                    del st.session_state['revision_data']
                st.session_state['revision_data'] = load_estimate_for_revision(search_rma)
        else:
            st.warning("Please enter an RMA number to search.")

    if 'revision_data' in st.session_state and st.session_state['revision_data'] is not None:
        data = st.session_state['revision_data']
        is_legacy = data.get('is_legacy', True)

        st.success(f"Found latest record for RMA: **{data.get('RMA')}**")

        if is_legacy:
            st.warning("This is a legacy estimate. Parts must be re-uploaded.", icon="⚠️")
            if st.button("Load Customer Details & Prepare for Upload", on_click=populate_form_for_revision, type="primary"):
                pass # The on_click handles the logic
        else:
            st.button("Load Full Estimate for Revision", on_click=populate_form_for_revision, type="primary")

    elif 'revision_data' in st.session_state and st.session_state['revision_data'] is None:
         st.error(f"No estimate found with RMA starting with '{st.session_state.search_rma_input}'.")

# Load the customer list into the session state once
if 'customer_list' not in st.session_state:
    st.session_state['customer_list'] = load_customer_list()

# --- Step 1: Customer Details & Parts Upload ---
st.subheader("1. Fill Details & Upload Parts Sheet")

col1, col2 = st.columns(2)
with col1:
    rma = st.text_input("RMA No.", value=st.session_state.get("rma", ""), key="rma")
    
    
    # UPDATE THE LINE BELOW
    cust_num = st.text_input(
        "Customer Number",
        value=st.session_state.get("cust_num", ""),
        key="cust_num",
        on_change=autofill_customer_name # <-- ADD THIS to trigger the autofill
    )
    cust_name = st.text_input("Customer Name", value=st.session_state.get("cust_name", ""), key="cust_name")
with col2:
    serial = st.text_input("Serial No.", value=st.session_state.get("serial", ""), key="serial")
    contact = st.text_input("Service Contact", value=st.session_state.get("contact", ""), key="contact")

description = st.text_area("Customer Description of Problem", value=st.session_state.get("description", ""), key="description", height=100)
evaluation = st.text_area("Technician Product Evaluation", value=st.session_state.get("evaluation", ""), key="evaluation", height=100)


st.subheader("Shipping Cost Calculator")

# Load shipping data once
zone_ranges_df = load_zone_ranges() # <-- Updated function call
shipping_prices_df = load_shipping_prices()

ship_col1, ship_col2, ship_col3 = st.columns([1, 1, 2])

with ship_col1:
    shipping_zip = st.text_input(
    "Shipping ZIP Code",
    value=st.session_state.get("shipping_zip", ""), # <-- We are adding this 'value' parameter
    key="shipping_zip"
    )

with ship_col2:
    item_type = st.selectbox(
        "Item Type for Shipping",
        options=["", "Laser console", "Delivery Device"],
        key="shipping_item_type"
    )

# --- New Calculation Logic for Ranges ---
if shipping_zip and item_type and not zone_ranges_df.empty and not shipping_prices_df.empty:
    if len(shipping_zip) >= 3 and shipping_zip.isdigit():
        zip_prefix_int = int(shipping_zip[:3])
        found_zone = None

        # Find which range the ZIP prefix falls into
        for _, row in zone_ranges_df.iterrows():
            if row['Start ZIP'] <= zip_prefix_int <= row['End ZIP']:
                found_zone = row['Zone']
                break # Found the correct range, stop searching
        
        if found_zone:
            try:
                price = shipping_prices_df.loc[str(found_zone), item_type]
                st.session_state['shipping_cost'] = float(price)
                with ship_col3:
                    st.metric("Calculated Shipping Cost", f"${float(price):,.2f}")
            except (KeyError, ValueError):
                st.session_state['shipping_cost'] = 0
                with ship_col3:
                    st.warning(f"No price found for Zone {found_zone} and item '{item_type}'.")
        else:
            st.session_state['shipping_cost'] = 0
            with ship_col3:
                st.error(f"Zone not found for ZIP prefix {zip_prefix_int}.")
    else:
        st.session_state['shipping_cost'] = 0
        if shipping_zip: # Only show warning if there's some input
            with ship_col3:
                 st.warning("Please enter a valid numeric ZIP code.")

st.markdown("---")

uploader_key = st.session_state.get('uploader_key', 0)
uploaded_file = st.file_uploader(
    "Choose the parts Excel file", 
    type="xlsx", 
    key=f"file_uploader_{uploader_key}" # This dynamic key is essential
)




# --- Step 2: Cost Preview (Activates after file upload) ---
if uploaded_file is not None or 'parts_df' in st.session_state:
    st.subheader("2. Review and Edit Estimated Cost")
    try:
        if uploaded_file is not None:
            # Store the bytes in session state for archiving later
            st.session_state['uploaded_file_bytes'] = uploaded_file.getvalue() 
            # Read from the stored bytes to create the DataFrame
            full_parts_df = pd.read_excel(BytesIO(st.session_state['uploaded_file_bytes']))
        else:
            full_parts_df = st.session_state['parts_df']

        # This ensures the column is always a float type, ready for decimals.
        full_parts_df['Amount Including Tax'] = full_parts_df['Amount Including Tax'].astype(float)
        
        if 'shipping_cost' in st.session_state and st.session_state['shipping_cost'] > 0:
            freight_mask = full_parts_df['No.'] == 'BILLABLE FREIGHT'
            if freight_mask.any():
                full_parts_df.loc[freight_mask, 'Amount Including Tax'] = st.session_state['shipping_cost']
                st.toast(f"Automatically set BILLABLE FREIGHT cost to ${st.session_state['shipping_cost']:,.2f}", icon="🚚")

        price_library = load_price_library()
        if price_library:
            autofill_count = 0
            correction_count = 0
            for index, row in full_parts_df.iterrows():
                part_no = str(row['No.'])
                if part_no in price_library and price_library[part_no] is not None:
                    library_price = float(price_library[part_no])
                    
                    if pd.isna(row['Amount Including Tax']) or row['Amount Including Tax'] == 0:
                        full_parts_df.at[index, 'Amount Including Tax'] = library_price
                        autofill_count += 1
                    else:
                        uploaded_price = float(row['Amount Including Tax'])
                        if uploaded_price > library_price:
                            full_parts_df.at[index, 'Amount Including Tax'] = library_price
                            correction_count += 1
                            st.toast(f"Price for {part_no} corrected to ${library_price:,.2f}", icon="✅")
            
            if autofill_count > 0:
                st.toast(f"Auto-filled {autofill_count} price(s) from the library.", icon="✨")
            if correction_count > 0:
                st.info(f"Applied {correction_count} price correction(s) to match the library.", icon="🛡️")

        technician_hq_mask = full_parts_df['No.'] == 'TECHNICIAN HQ'
        if technician_hq_mask.any():
            full_parts_df.loc[technician_hq_mask, 'Amount Including Tax'] = 285

        if 'Amount Including Tax' not in full_parts_df.columns:
            st.error("The uploaded file is missing the 'Amount Including Tax' column.")
            st.stop()
        if 'Quantity' not in full_parts_df.columns:
            st.error("The uploaded file is missing the 'Quantity' column.")
            st.stop()

        full_parts_df['Amount Including Tax'] = pd.to_numeric(full_parts_df['Amount Including Tax'], errors='coerce').fillna(0)
        full_parts_df['Quantity'] = pd.to_numeric(full_parts_df['Quantity'], errors='coerce').fillna(1)

        columns_to_display = ["No.", "Description", "Quantity", "Amount Including Tax"]
        display_df = full_parts_df[columns_to_display]

        st.info("Edit the unit prices in the 'Amount Including Tax' column. The total cost will update automatically.")

        edited_display_df = st.data_editor(
            display_df,
            num_rows="dynamic",
            column_config={
                "Amount Including Tax": st.column_config.NumberColumn(
                    "Unit Price (incl. Tax)",
                    format="$%.2f",
                    step=0.01,
                )
            },
            key="parts_editor"
        )

        full_parts_df.set_index('No.', inplace=True)
        edited_display_df.set_index('No.', inplace=True)
        full_parts_df.update(edited_display_df)
        full_parts_df.reset_index(inplace=True)

        st.session_state['parts_for_library'] = full_parts_df.copy()

        full_parts_df['Line Total'] = full_parts_df['Quantity'] * full_parts_df['Amount Including Tax']

        total_cost = full_parts_df['Line Total'].sum()
        st.metric("Total Estimated Cost", f"${total_cost:,.2f}")

        final_estimate_df = full_parts_df.copy()
        final_estimate_df['Amount Including Tax'] = final_estimate_df['Line Total']

        st.session_state['parts_df'] = final_estimate_df

    except Exception as e:
        st.error(f"Error reading or processing the Excel file: {e}")
        st.stop()

# --- Step 3: Generation Form (Activates after cost is previewed) ---
if 'parts_df' in st.session_state:
    with st.form(key="generation_form"):
        # ...
        submitted = st.form_submit_button("Generate & Log Estimate to 'Estimate Form MOAS'")

    if submitted:
        if not rma or not cust_name:
            st.error("⚠️ Please ensure RMA No. and Customer Name are filled out.")
        else:
            with st.spinner("Processing... Please wait."):
                try:
                    # --- NEW LOGIC: Archive the source parts file ---
                    if 'uploaded_file_bytes' in st.session_state:
                        if not os.path.exists(SOURCE_PARTS_ARCHIVE_DIR):
                            os.makedirs(SOURCE_PARTS_ARCHIVE_DIR)
                        
                        archive_path = os.path.join(SOURCE_PARTS_ARCHIVE_DIR, f"{rma}.xlsx")
                        with open(archive_path, "wb") as f:
                            f.write(st.session_state['uploaded_file_bytes'])
                        
                        # Clean up to prevent re-saving on refresh
                        del st.session_state['uploaded_file_bytes']
                    parts_df_for_generation = st.session_state['parts_df']
                    parts_df_for_library = st.session_state['parts_for_library']
                    form_data = {key: val for key, val in st.session_state.items()}

                    file_paths = generate_estimate_files(form_data, parts_df_for_generation, SAVE_DIRECTORY)

                    if file_paths:
                        st.success("✅ Documents generated successfully!")
                        st.session_state['file_paths'] = file_paths

                        with st.spinner("Updating 'Estimate Form MOAS' sheet..."):
                            if add_or_update_estimate_in_gsheet(form_data, parts_df_for_generation):
                                st.success("✅ Estimate logged successfully.")
                                with st.spinner("Updating Price Library and Usage Counts..."):
                                    update_price_library_and_usage_count(parts_df_for_library)
                                st.cache_data.clear()
                            else:
                                st.error("❌ Failed to log estimate.")
                    else:
                        st.error("❌ There was an error generating the documents.")
                except Exception as e:
                    st.error(f"An unexpected error occurred: {e}")

# --- Step 4: Action Buttons (Appear after successful generation) ---
#if 'file_paths' in st.session_state:
 #   st.subheader("4. Download or Email")
if 'file_paths' in st.session_state:
    st.subheader("4. Download Generated Estimate")

    excel_path = st.session_state['file_paths']['excel_path']

    if excel_path:
        st.info("The estimate has been generated as an Excel file. You can now download it.")
        with open(excel_path, "rb") as excel_file:
            st.download_button(
                label="Download Estimate Excel File",
                data=excel_file,
                file_name=os.path.basename(excel_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    st.markdown("---")
    
    pdf_path = st.session_state['file_paths']['pdf_path']
    rma_on_form = st.session_state.get('rma', '')
    serial_on_form = st.session_state.get('serial', '')

    with open(pdf_path, "rb") as pdf_file:
        st.download_button(
            label="Download Estimate PDF",
            data=pdf_file,
            file_name=os.path.basename(pdf_path),
            mime="application/pdf",
            key="download_estimate_pdf"
        )

    st.write("---")
    st.subheader("Send Estimate and Update S/N Email History")

    rma_for_history = st.text_input(
        "Enter the Service Request number that connects to the RMA'",
        value=st.session_state.get("rma_for_history", rma_on_form),
        key="rma_for_history",
        help="Enter the Service Request number that corresponds to the record you want to update in the S/N EMAIL history sheet."
    )

    recipient = st.text_input("Recipient Email:", st.session_state.get('', ''))

    if st.button("📧 Send Email & Update Log"):
        if recipient and "@" in recipient and rma_for_history:
            with st.spinner("Sending email... Please check Outlook for progress."):
                email_success, cc_form_path = send_estimate_email(recipient, rma_on_form, serial_on_form, pdf_path)

                if email_success:
                    st.success(f"Email sent to {recipient}!")
                    st.session_state['cc_form_path'] = cc_form_path

                    if update_estimate_sent_details_in_gsheet(rma_for_history, st.session_state.get('serial'), recipient, date.today()):
                         st.info(f"Marked as 'Estimate Sent' for RMA {rma_for_history} in S/N EMAIL history.")
                         st.cache_data.clear()
                    else:
                        st.error(f"Could not find or update RMA {rma_for_history} in the S/N EMAIL history sheet.")
                else:
                    st.error("Failed to send email. Ensure Outlook is running.")
        else:
            st.warning("Please enter a valid Recipient Email and the RMA to update in the history log.")

    if 'cc_form_path' in st.session_state and st.session_state['cc_form_path']:
        cc_path = st.session_state['cc_form_path']
        if os.path.exists(cc_path):
            with open(cc_path, "rb") as cc_file:
                st.download_button(
                    label="Download Credit Card Form PDF",
                    data=cc_file,
                    file_name=os.path.basename(cc_path),
                    mime="application/pdf",
                    key="download_cc_pdf"
                )
    st.markdown("---")
    st.button("Clear Form", on_click=clear_form_state, type="secondary")
