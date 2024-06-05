# streamlit_app.py

import streamlit as st
import pandas as pd
from io import BytesIO

# Function to process the data
def process_excel(file):
    try:
        # Read the "Sponsored Products Campaigns" sheet into a DataFrame
        sp_df = pd.read_excel(file, sheet_name='Sponsored Products Campaigns')

        # Filter the DataFrame to only include rows where 'Entity' is 'Keyword'
        filtered_sp_df = sp_df[sp_df['Entity'] == 'Keyword']

        # Further filter the DataFrame to include only rows where 'Units' is 0 and 'Spend' is greater than 0
        filtered_sp_df = filtered_sp_df[(filtered_sp_df['Units'] == 0) & (filtered_sp_df['Spend'] > 0)]

        # Set 'Operation' column to 'Update' and 'State' column to 'paused'
        filtered_sp_df['Operation'] = 'Update'
        filtered_sp_df['State'] = 'paused'

        return filtered_sp_df
    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
        return None

# Function to convert DataFrame to Excel for download
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Filtered Keywords')
    processed_data = output.getvalue()
    return processed_data

# Streamlit app
st.title("SP Keywords to Pause")
st.write("Script pauses all spending keywords with zero sales.")
st.write("Input file - Amazon bulk file, timeframe - 8 weeks.")

# File uploader
uploaded_file = st.file_uploader("Upload Amazon Bulk File", type="xlsx")

if uploaded_file is not None:
    # Process the uploaded file
    result_df = process_excel(uploaded_file)
    
    if result_df is not None:
        # Display the first few rows of the result
        st.write("SP Keywords to Pause")
        st.dataframe(result_df.head())
        
        # Provide download link for the processed file
        st.write("Download the processed file:")
        st.download_button(
            label="Download Excel file",
            data=to_excel(result_df),
            file_name="Keywords to pause.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
