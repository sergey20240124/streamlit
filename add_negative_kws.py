import streamlit as st
import pandas as pd
from io import BytesIO

# Function to process the data
def process_excel(file, additional_spend, additional_acos):
    try:
        # Read the "Sponsored Products Campaigns" sheet into a DataFrame
        sp_df = pd.read_excel(file, sheet_name='Sponsored Products Campaigns')

        # Filter the DataFrame to only include rows where 'Entity' is 'Keyword'
        filtered_sp_df = sp_df[sp_df['Entity'] == 'Keyword']

        # Automatically extract keywords with Units == 0 and Spend > 0
        auto_filtered_sp_df = filtered_sp_df[(filtered_sp_df['Units'] == 0) & 
                                             (filtered_sp_df['Spend'] > 0)]

        # Convert ACOS from percentage to a numerical value
        additional_acos /= 100.0

        # Apply additional filters based on user input: Spend > additional_spend and ACOS > additional_acos
        if additional_spend > 0 or additional_acos > 0:
            additional_filtered_sp_df = filtered_sp_df[(filtered_sp_df['Units'] > 0) & 
                                                       (filtered_sp_df['Spend'] > additional_spend) & 
                                                       (filtered_sp_df['ACOS'] > additional_acos)]
            # Combine both filtered DataFrames
            combined_filtered_df = pd.concat([auto_filtered_sp_df, additional_filtered_sp_df])
        else:
            combined_filtered_df = auto_filtered_sp_df

        # Remove duplicates
        combined_filtered_df = combined_filtered_df.drop_duplicates()

        # Set 'Operation' column to 'Update' and 'State' column to 'paused'
        combined_filtered_df['Operation'] = 'Update'
        combined_filtered_df['State'] = 'paused'

        return combined_filtered_df
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
st.write("""
This app automatically extracts spending keywords with zero sales.<br>
Upon processing, the app creates an output file ready for uploading via Ad Console.<br><br>

Input file - Bulk Report XLSX <br>
Required sheets - SP Campaigns<br><br>

For more sophisticated scenarios, use the inputs below to select additional keywords to pause.<br><br>

""", unsafe_allow_html=True)

# Get additional filter inputs from the user
additional_spend_input = st.number_input("Minimum Spend", min_value=0.0, value=0.0)
additional_acos_input = st.number_input("Target ACOS (%)", min_value=0.0, value=0.0)

# File uploader
uploaded_file = st.file_uploader("Upload Amazon Bulk File", type="xlsx")

if uploaded_file is not None:
    # Process the uploaded file
    result_df = process_excel(uploaded_file, additional_spend_input, additional_acos_input)
    
    if result_df is not None and not result_df.empty:
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
    else:
        st.write("No keywords matched the specified filters.")
