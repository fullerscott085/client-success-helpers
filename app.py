import streamlit as st
import pandas as pd
import io
from main import dataframes_for_export, KeyItemCollection, KeyItem, DataType, process_zip_archive

uploaded_file = st.file_uploader(
    ".zip file containing CI's", 
    type=['.zip'], 
    accept_multiple_files=False, 
    key='ci-files', 
    help='A zip file containing all the Oracle Commercial invoices. Each file should be a PDF with exactly the same format.'
)

collection = KeyItemCollection([
        KeyItem(key='gross-weight', search_text='Gross Weight', data_type=DataType.TEXT),
        KeyItem(key='comm-inv-no', search_text='Comm Inv No', data_type=DataType.TEXT)
    ])

if uploaded_file is not None:
    results = process_zip_archive(uploaded_file, collection)
    df_for_salesforce, df = dataframes_for_export(results)
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_for_salesforce.to_excel(writer, sheet_name='BOM Cals', index=False)
        df.to_excel(writer, sheet_name='PDF Data', index=False)

    # Ensure buffer is ready for download
    output.seek(0)

    st.download_button(
        label="Download Excel",
        data=output,
        file_name="results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon="⬇️"
    )
