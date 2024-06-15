import pandas as pd
import streamlit as st
from io import BytesIO

# Define functions
def filter_enabled(df, column_names):
    filtered_df = df.copy()
    for column in column_names:
        if column in df.columns:
            filtered_df = filtered_df[filtered_df[column].str.lower() == 'enabled']
    filtered_df.reset_index(drop=True, inplace=True)
    return filtered_df

def filter_keywords(df, column_names):
    keywords_only_df = df.copy()
    for column in column_names:
        if column in df.columns:
            keywords_only_df = keywords_only_df[keywords_only_df[column].str.lower() == 'keyword']
    keywords_only_df.reset_index(drop=True, inplace=True)
    return keywords_only_df

def create_comparison_df(sp_performing_keywords_only_df, sb_keywords_only_df):
    sp_keywords = set(sp_performing_keywords_only_df['Keyword Text'])
    sb_keywords = set(sb_keywords_only_df['Keyword Text'])
    unique_sp_keywords = sp_keywords - sb_keywords

    # Convert sets to lists and ensure they are of the same length
    sp_keywords_list = list(sp_keywords)
    sb_keywords_list = list(sb_keywords)
    unique_sp_keywords_list = list(unique_sp_keywords)

    max_length = max(len(sp_keywords_list), len(sb_keywords_list), len(unique_sp_keywords_list))

    sp_keywords_list.extend([''] * (max_length - len(sp_keywords_list)))
    sb_keywords_list.extend([''] * (max_length - len(sb_keywords_list)))
    unique_sp_keywords_list.extend([''] * (max_length - len(unique_sp_keywords_list)))

    # Create DataFrame for export
    df_export = pd.DataFrame({
        'Performing SP Keywords': sp_keywords_list,
        'Existing SB Keywords': sb_keywords_list,
        'Keyword to consider adding to SB campaign': unique_sp_keywords_list
    })

    return df_export

# Streamlit app
st.title('Growth Opportunities: Keywords')

# Input target ACOS
target_acos = st.number_input('Enter target ACOS:', min_value=0.0, max_value=1.0, step=0.01, value=0.25)

# File upload
uploaded_file = st.file_uploader('Upload the bulk file', type=['xlsx'])

if uploaded_file:
    try:
        # Read the uploaded Excel file
        sp_df = pd.read_excel(uploaded_file, sheet_name='Sponsored Products Campaigns')
        sb_df = pd.read_excel(uploaded_file, sheet_name='Sponsored Brands Campaigns')

        # Convert 'Units' and 'ACOS' columns to numeric, handling errors gracefully
        sp_df['Units'] = pd.to_numeric(sp_df['Units'], errors='coerce')
        sp_df['ACOS'] = pd.to_numeric(sp_df['ACOS'], errors='coerce')

        # Check for NaN values after conversion and drop rows with NaNs in 'Units' or 'ACOS'
        sp_df = sp_df.dropna(subset=['Units', 'ACOS'])

        # Filter 'enabled' rows for both DataFrames
        filtered_sp_df = filter_enabled(sp_df, ['State', 'Campaign State (Informational only)', 'Ad Group State (Informational only)'])
        filtered_sb_df = filter_enabled(sb_df, ['State', 'Campaign State (Informational only)'])

        # Ensure 'Keyword Text' column is treated as strings and handle NaNs
        filtered_sp_df['Keyword Text'] = filtered_sp_df['Keyword Text'].astype(str).fillna('')

        # Filter out keywords containing '+'
        filtered_sp_df = filtered_sp_df[~filtered_sp_df['Keyword Text'].str.contains('\+')]

        # Apply the filter_keywords function to filtered DataFrames
        sp_keywords_only_df = filter_keywords(filtered_sp_df, ['Entity'])
        sb_keywords_only_df = filter_keywords(filtered_sb_df, ['Entity'])

        # Filtering performing keywords based on 'Units' and 'ACOS'
        sp_performing_keywords_only_df = sp_keywords_only_df[(sp_keywords_only_df['Units'] >= 2) & (sp_keywords_only_df['ACOS'] < target_acos)]
        sp_performing_keywords_only_df.reset_index(drop=True, inplace=True)

        # Create comparison DataFrame
        df_export = create_comparison_df(sp_performing_keywords_only_df, sb_keywords_only_df)

        # Convert DataFrame to Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Keywords Comparison')
        processed_data = output.getvalue()

        # Download button
        st.download_button(label='Download comparison file', data=processed_data, file_name='keywords_comparison.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

       
    except Exception as e:
        st.error(f"An error occurred: {e}")
