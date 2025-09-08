# pages/2_📚_Bulk_Update_Library.py

import streamlit as st
import pandas as pd
from logic import process_historical_usage

st.set_page_config(layout="wide")
st.title("📊 Process Historical Usage by Quarter")
st.info(
    "Upload multiple historical estimate Excel files here. The application will read each file, "
    "extract the parts and the 'Planned Delivery Date', and update the Price Library with usage "
    "counts broken down by year and quarter."
)
st.warning(
    "This process will overwrite the existing Price Library to add new quarterly columns. "
    "Ensure your files contain a 'Planned Delivery Date' column.",
    icon="⚠️"
)

uploaded_files = st.file_uploader(
    "Select one or more historical estimate Excel files",
    accept_multiple_files=True,
    type="xlsx",
    key="bulk_uploader_quarterly"
)

if uploaded_files:
    st.markdown("---")
    
    if st.button("Process Files and Update Quarterly Counts", type="primary"):
        with st.spinner("Processing files and aggregating data... Please wait."):
            all_parts_list = []
            files_processed = 0
            files_with_errors = []

            for uploaded_file in uploaded_files:
                try:
                    df = pd.read_excel(uploaded_file)
                    
                    # Basic check for required columns
                    required_cols = {'No.', 'Description', 'Planned Delivery Date', 'Amount Including Tax', 'Quantity'}
                    if not required_cols.issubset(df.columns):
                        st.warning(f"Skipping `{uploaded_file.name}`: It's missing one or more required columns.")
                        files_with_errors.append(uploaded_file.name)
                        continue

                    all_parts_list.append(df)
                    files_processed += 1

                except Exception as e:
                    st.warning(f"Could not process file `{uploaded_file.name}`. Error: {e}")
                    files_with_errors.append(uploaded_file.name)

            if not all_parts_list:
                st.error("No valid data could be extracted from the uploaded files.")
            else:
                # Create one master table with all parts from all files
                master_df = pd.concat(all_parts_list, ignore_index=True)
                
                # The new logic function does all the heavy lifting
                update_success = process_historical_usage(master_df)

                if update_success:
                    st.success("✅ Quarterly usage counts have been successfully updated in the Price Library!")
                    st.cache_data.clear() # Clear cache to reflect changes immediately
                else:
                    st.error("❌ A failure occurred while processing historical data.")
            
            if files_with_errors:
                st.warning(f"The following files could not be processed: {', '.join(files_with_errors)}")
