import streamlit as st
import pandas as pd
from io import BytesIO
import warnings

# Suppress the specific UserWarning from openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Data cleaning function
def clean_amazon_data(file):
    amazon_data = pd.read_excel(file)
    amazon_data.columns = amazon_data.columns.str.strip()
    amazon_data['Date'] = pd.to_datetime(amazon_data['Date'])
    amazon_data.fillna(0, inplace=True)

    columns_to_remove = [
        'Top-of-search Impression Share', 'Total Advertising Cost of Sales (ACOS)',
        'Total Return on Advertising Spend (ROAS)', '7 Day Total Orders (#)',
        '7 Day Conversion Rate', '7 Day Advertised SKU Units (#)',
        '7 Day Other SKU Units (#)', '7 Day Advertised SKU Sales', '7 Day Other SKU Sales',
        'Currency', 'Ad Group Name'
    ]
    amazon_data.drop(columns=columns_to_remove, inplace=True)

    amazon_data.rename(columns={
        '7 Day Total Sales': 'Ad Sales', 
        'Spend': 'Ad Spend', 
        '7 Day Total Units (#)': 'Units',
        'Cost Per Click (CPC)': 'CPC',
        'Click-Thru Rate (CTR)': 'CTR',
        'Portfolio name': 'Portfolio'
    }, inplace=True)

    numeric_columns = ['Impressions', 'Ad Sales', 'Units']
    amazon_data[numeric_columns] = amazon_data[numeric_columns].apply(pd.to_numeric, errors='coerce')

    amazon_data['CPC'] = amazon_data['CPC'].round(2)
    amazon_data['CTR'] = amazon_data['CTR'].round(2)

    amazon_data = amazon_data[amazon_data['Ad Spend'] > 0]

    exclude_targeting = ['loose-match', 'close-match', 'complements', 'substitutes']
    amazon_data = amazon_data[~amazon_data['Targeting'].isin(exclude_targeting)]

    exclude_keywords = ['category', 'B0']
    pattern = '|'.join(exclude_keywords)
    amazon_data = amazon_data[~amazon_data['Targeting'].str.contains(pattern, case=False, na=False)]

    amazon_data.reset_index(drop=True, inplace=True)

    return amazon_data

# Function to convert DataFrame to Excel in memory
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()  # Corrected method
    processed_data = output.getvalue()
    return processed_data

# Streamlit app
def main():
    st.title("Amazon Data Cleaner")

    uploaded_file = st.file_uploader("Choose an XLSX file", type="xlsx")

    if uploaded_file is not None:
        st.write("File uploaded successfully!")
        
        # Clean the data
        cleaned_data = clean_amazon_data(uploaded_file)
        
        # Display a preview of the cleaned data
        st.write("Preview of cleaned data:")
        st.dataframe(cleaned_data.head())

        # Convert the cleaned data to an Excel file in memory
        cleaned_data_excel = to_excel(cleaned_data)

        # Allow the user to download the cleaned data
        st.download_button(
            label="Download cleaned data as XLSX",
            data=cleaned_data_excel,
            file_name='cleaned_data.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == "__main__":
    main()
