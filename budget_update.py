import streamlit as st
import pandas as pd
import warnings
from io import BytesIO

# Suppress the specific UserWarning from openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Function to process the data
def process_data(file):
    try:
        # Load the Excel file
        sp_df = pd.read_excel(file, sheet_name='Sponsored Products Campaigns')

        # Filter sp_df for 'enabled'
        sp_df = sp_df[(sp_df['State'] == 'enabled') & (sp_df['Campaign State (Informational only)'] == 'enabled')]
        sp_df = sp_df.reset_index(drop=True)  # Reindex sp_df

        # Drop duplicates
        sp_df.drop_duplicates(inplace=True)

        # Fill missing values
        for col in sp_df.columns:
            if sp_df[col].dtype == 'object':
                sp_df[col] = sp_df[col].fillna('')  # Fill text columns with empty string
            else:
                sp_df[col] = sp_df[col].fillna(0)   # Fill numeric columns with 0

        # Trim whitespaces
        text_cols = sp_df.select_dtypes(include=['object']).columns
        for col in text_cols:
            sp_df[col] = sp_df[col].str.strip()

        # Filter Entity for Campaign only
        campaign_sp_df = sp_df[sp_df['Entity'] == 'Campaign']
        campaign_sp_df = campaign_sp_df.reset_index(drop=True)

        # Create 'POB' column with values Spend/14/Daily Budget
        campaign_sp_df['POB'] = campaign_sp_df.apply(lambda row: row['Spend'] / 14 / row['Daily Budget'] if row['Daily Budget'] != 0 else 0, axis=1)

        # Create 'Bud Ref' column with values mirroring Daily Budget
        campaign_sp_df['Bud Ref'] = campaign_sp_df['Daily Budget']

        # Filter 80% camps for Campaign only
        selected_sp_df = campaign_sp_df[campaign_sp_df['POB'] >= 0.8]
        selected_sp_df = selected_sp_df.reset_index(drop=True)

        # Update Daily Budget column values to Bud Ref * 1.2
        selected_sp_df['Daily Budget'] = selected_sp_df.apply(lambda row: row['Bud Ref'] * 1.2 if row['Bud Ref'] != 0 else 0, axis=1)
        selected_sp_df = selected_sp_df.drop(columns=['POB', 'Bud Ref'])

        return selected_sp_df
    except Exception as e:
        st.error(f"Error processing data: {e}")
        return pd.DataFrame()  # Return an empty DataFrame on error

# Streamlit app
st.title("Budget Review")

# File uploader
uploaded_file = st.file_uploader("Choose a Bulk file", type="xlsx")

if uploaded_file is not None:
    # Process the uploaded file
    selected_sp_df = process_data(uploaded_file)

    if not selected_sp_df.empty:
       
        # Download the processed DataFrame as an Excel file
        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            writer.close()
            processed_data = output.getvalue()
            return processed_data

        processed_data = to_excel(selected_sp_df)

        st.download_button(label="Download Processed Data as Excel",
                           data=processed_data,
                           file_name='Campaign Bydgets Updated.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        st.error("No data to display. Please check the input file and try again.")
