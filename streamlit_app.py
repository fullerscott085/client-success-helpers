import streamlit as st
import pandas as pd
import datetime
import io
from main import dataframes_for_export, KeyItemCollection, KeyItem, DataType, process_zip_archive

st.title("Oracle Commercial Invoice PDF Processor")

# Upload ZIP File
uploaded_file = st.file_uploader(
    ".zip file containing CI's", 
    type=['zip'], 
    accept_multiple_files=False, 
    key='ci-files', 
    help="A ZIP file containing Oracle Commercial invoices. Each file should be a PDF with exactly the same format."
)

if uploaded_file is not None:
    # Save the file to disk
    with open(f"./uploaded_files/{uploaded_file.name}", "wb") as f:
        f.write(uploaded_file.getbuffer())
    print(f"File saved as {uploaded_file.name}")

if uploaded_file is not None and "processed_data" not in st.session_state:
    print(f"\nSTART-TIME: {datetime.datetime.now().isoformat()}\n")
    status_text = st.empty()  # Placeholder for status updates
    progress_bar = st.progress(0)  # Progress bar

    status_text.text("📂 Upload successful. Processing started...")

    # Define search keys
    collection = KeyItemCollection([
        KeyItem(key='gross-weight', search_text='Gross Weight', data_type=DataType.TEXT),
        KeyItem(key='comm-inv-no', search_text='Comm Inv No', data_type=DataType.TEXT)
    ])

    # Callback function to update progress
    def update_progress(current, total):
        # print(f"current: {current}")
        # print(f"total: {total}")
        progress = int((current / total) * 100)
        # print(f"progress: {progress}")
        progress_bar.progress(progress)
        status_text.text(f"🔍 Processing file {current} of {total}...")

    # Process the ZIP file
    results = process_zip_archive(
        zip_path=uploaded_file, 
        collection=collection, 
        progress_callback=update_progress
    )
    status_text.text(f"🔍 Extracted data from {len(results)} PDFs...")

    # Convert results to DataFrames
    df_for_salesforce, df = dataframes_for_export(results)
    status_text.text("📊 Data processing complete. Preparing Excel file...")
    
    # Create Excel file in memory   
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_for_salesforce.to_excel(writer, sheet_name='BOM Cals', index=False)
        df.to_excel(writer, sheet_name='PDF Data', index=False)

    # Ensure buffer is ready for download
    output.seek(0)
    
    # Store results in session state to persist across reruns
    # st.session_state["processed_data"] = results
    # st.session_state["excel_file"] = output.getvalue()
    
    status_text.text("📝 Excel file created. Finalizing...")
    status_text.text("✅ Process complete. Download your Excel file below.")

    st.download_button(
        label="Download Excel",
        data=output.getvalue(), #st.session_state['excel_file'],
        file_name="results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon="⬇️"
    )
    print(f"\nEND-TIME: {datetime.datetime.now().isoformat()}\n")
