import streamlit as st
import pandas as pd
import warnings
from datetime import datetime
import io

# Suppress the specific UserWarning from openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

st.title("SP Performing Search Terms")
st.write('''
This app extracts Sponsored Products performing customer search terms.<br>
Upon processing app creates output file ready for uploading via Ad Console.<br>
Input file - &nbsp;Bulk Report XLSX <br>
Required sheets - &nbsp;Portfolios, SP Campaigns, SP Search terms")<br>
Extraction criteria - &nbsp;Units>=2, below Target ACOS")<br>
Output file - &nbsp;Bulk XLSX
"", unsafe_allow_html=True)


      

# ACOS input
acos_input = st.number_input("Enter target ACOS (%)", min_value=0.0, max_value=100.0, value=15.0)
acos_value = acos_input / 100  # Convert to decimal for calculations

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Load the Excel file
    file_path = uploaded_file

    # READ BULK SHEET
    portfolios_df = pd.read_excel(file_path, sheet_name='Portfolios')
    sp_df = pd.read_excel(file_path, sheet_name='Sponsored Products Campaigns')
    sts_sp_df = pd.read_excel(file_path, sheet_name='SP Search Term Report')

    # CLEAN AND MODIFY
    sp_df['Portfolio Name (Informational only)'] = sp_df['Portfolio Name (Informational only)'].fillna('No portfolio')
    sts_sp_df['Portfolio Name (Informational only)'] = sts_sp_df['Portfolio Name (Informational only)'].fillna('No portfolio')

    sp_df['Portfolio ID'] = None

    def fill_portfolio_id(row):
        if row['Portfolio Name (Informational only)'] == 'No portfolio':
            return ''
        else:
            portfolio_row = portfolios_df[portfolios_df['Portfolio Name'].str.strip() == row['Portfolio Name (Informational only)'].strip()]
            if not portfolio_row.empty:
                return str(portfolio_row['Portfolio ID'].values[0])
        return row['Portfolio ID']

    sp_df['Portfolio ID'] = sp_df.apply(fill_portfolio_id, axis=1)
    sts_sp_df['Portfolio ID'] = ''
    sts_sp_df['Portfolio ID'] = sts_sp_df.apply(fill_portfolio_id, axis=1)

    sp_df.columns = sp_df.columns.str.strip()

    # Exclude rows where Entity is 'Product Targeting'
    sp_df = sp_df[sp_df['Entity'] != 'Product Targeting']

    # CALCULATING BIDS AND BUDGETS
    avg_cpc_df = sts_sp_df.groupby(['Customer Search Term', 'Portfolio ID'])['CPC'].mean().reset_index()
    avg_cpc_df = avg_cpc_df.rename(columns={'Customer Search Term': 'Keyword', 'CPC': 'Avg CPC'})

    sp_df['Daily Budget'] = pd.to_numeric(sp_df['Daily Budget'], errors='coerce')
    avg_budget_df = sp_df.groupby('Portfolio ID')['Daily Budget'].mean().reset_index()
    avg_budget_df = avg_budget_df.rename(columns={'Daily Budget': 'Avg Daily Budget'})

    # EXTRACTING BEST PERFORMING SKUS BY PORTFOLIO
    portfolio_columns = sp_df[['Portfolio ID', 'Portfolio Name (Informational only)', 'SKU', 'Sales']]
    portfolio_columns = portfolio_columns.dropna(subset=['SKU', 'Sales'])
    best_sku_per_portfolio = portfolio_columns.loc[portfolio_columns.groupby(['Portfolio ID', 'Portfolio Name (Informational only)'])['Sales'].idxmax()]
    best_sku_per_portfolio = best_sku_per_portfolio.rename(columns={'Portfolio Name (Informational only)': 'Portfolio Name', 'SKU': 'SKU'})
    best_sku_per_portfolio = best_sku_per_portfolio.reset_index(drop=True)
    
    # EXISTING KWS DF CREATED AND CLEANED
    existing_keywords_df = sp_df[sp_df['Entity'] == 'Keyword'][['Keyword Text', 'Campaign Name (Informational only)', 'Ad Group Name (Informational only)']].rename(columns={'Keyword Text': 'Keyword', 'Campaign Name (Informational only)': 'Campaign', 'Ad Group Name (Informational only)': 'Ad Group'})
    existing_keywords_df = existing_keywords_df.drop_duplicates(subset=['Keyword', 'Campaign', 'Ad Group'], keep='first')
    existing_keywords_df = existing_keywords_df[~existing_keywords_df['Keyword'].str.contains(r'\+')]
    existing_keywords_df.reset_index(drop=True, inplace=True)

    # PERFORMING SEARCH TERMS EXTRACTION AND CHECKING AGAINST EXISTING KEYWORDS
    performing_sts_df = sts_sp_df[(sts_sp_df['Units'] >= 2) & (sts_sp_df['ACOS'] <= acos_value) & (sts_sp_df['Product Targeting Expression'].isna())][['Customer Search Term', 'Campaign Name (Informational only)', 'Ad Group Name (Informational only)', 'Units', 'ACOS', 'Portfolio ID']]
    performing_sts_df = performing_sts_df.rename(columns={'Customer Search Term': 'Keyword', 'Campaign Name (Informational only)': 'Campaign', 'Ad Group Name (Informational only)': 'Ad Group'})
    performing_sts_df = performing_sts_df.drop_duplicates(subset=['Keyword', 'Campaign', 'Ad Group'])

    # Filter out search terms containing 'b0'
    performing_sts_df = performing_sts_df[~performing_sts_df['Keyword'].str.contains(r'b0', case=False, na=False)]

    # Identify duplicates with existing keywords
    duplicate_keywords_df = performing_sts_df.merge(existing_keywords_df, how='inner', on=['Keyword'])

    # Remove duplicates from performing_sts_df
    performing_sts_df = performing_sts_df[~performing_sts_df['Keyword'].isin(duplicate_keywords_df['Keyword'])]

    # Ensure no duplicates within performing_sts_df
    performing_sts_df = performing_sts_df.drop_duplicates(subset=['Keyword', 'Campaign', 'Ad Group'])
    performing_sts_df.reset_index(drop=True, inplace=True)    
    
    # Merge performing_sts_df with best_sku_per_portfolio to get the 'Best SKU' column
    performing_sts_df = performing_sts_df.merge(best_sku_per_portfolio[['Portfolio ID', 'SKU']], on='Portfolio ID', how='left')

    # Group keywords by SKU
    grouped_keywords_by_sku = performing_sts_df.groupby('SKU')['Keyword'].apply(list).reset_index()

    # Initialize the Portfolio ID column
    grouped_keywords_by_sku['Portfolio ID'] = None

    # Loop through each SKU group
    for index, row in grouped_keywords_by_sku.iterrows():
        sku = row['SKU']
        
        # Get the portfolios linked to this SKU's keywords in performing_sts_df
        linked_portfolios = performing_sts_df[performing_sts_df['SKU'] == sku]['Portfolio ID'].unique()
        
        # Get the portfolio with the highest average budget from avg_budget_df
        best_portfolio = avg_budget_df[avg_budget_df['Portfolio ID'].isin(linked_portfolios)].sort_values(by='Avg Daily Budget', ascending=False).iloc[0]['Portfolio ID']
        
        # Assign the best portfolio to the grouped_keywords_by_sku DataFrame
        grouped_keywords_by_sku.at[index, 'Portfolio ID'] = best_portfolio

    # Ensure Portfolio ID columns are of the same type (string)
    grouped_keywords_by_sku['Portfolio ID'] = grouped_keywords_by_sku['Portfolio ID'].astype(str)
    portfolios_df['Portfolio ID'] = portfolios_df['Portfolio ID'].astype(str)

    # Merge grouped_keywords_by_sku with portfolios_df to add the Portfolio Name column
    grouped_keywords_by_sku = grouped_keywords_by_sku.merge(portfolios_df[['Portfolio ID', 'Portfolio Name']], on='Portfolio ID', how='left')

    # FINAL PART

    # Initialize final_output_df with the specified columns
    final_output_df = pd.DataFrame(columns=[
        'Product', 'Entity', 'Operation', 'Campaign ID', 'Ad Group ID', 'Portfolio ID', 
        'Campaign Name', 'Ad Group Name', 'Start Date', 'End Date', 'Targeting Type', 
        'State', 'Daily Budget', 'SKU', 'Ad Group Default Bid', 'Bid', 'Keyword Text', 
        'Match Type', 'Bidding Strategy'
    ])

    # Fill the 'Product' column uniformly with 'Sponsored Products'
    final_output_df['Product'] = 'Sponsored Products'

    # Get the current date
    current_date = datetime.now().strftime('%Y-%m-%d')

    # Calculate the overall average CPC
    overall_avg_cpc = avg_cpc_df['Avg CPC'].mean().round(2)

    # Function to append a row to final_output_df, excluding all-NA rows
    def append_row(df, row_dict):
        row_df = pd.DataFrame([row_dict])
        # Exclude if the row is all-NA
        if not row_df.isnull().all(axis=1).all():
            return pd.concat([df, row_df], ignore_index=True)
        return df

    # Iterate through each row in grouped_keywords_by_sku
    for index, row in grouped_keywords_by_sku.iterrows():
        portfolio_name = row['Portfolio Name']
        sku = row['SKU']
        keywords = row['Keyword']
        portfolio_id = row['Portfolio ID']
        
        # Create the campaign name and ID using the naming convention
        campaign_name = f"{portfolio_name} | SP | {sku} | EXACT | STA | {current_date}"
        
        # Retrieve the Average Daily Budget
        avg_daily_budget_series = avg_budget_df.loc[avg_budget_df['Portfolio ID'] == portfolio_id, 'Avg Daily Budget']
        if not avg_daily_budget_series.empty:
            daily_budget = round(avg_daily_budget_series.values[0], 2)
        else:
            daily_budget = overall_avg_cpc  # Use overall average CPC as default daily budget
        
        # Append Campaign row
        campaign_row = {
            'Product': 'Sponsored Products',
            'Entity': 'Campaign',
            'Operation': 'Create',
            'Campaign ID': campaign_name,
            'Ad Group ID': '',
            'Portfolio ID': portfolio_id,
            'Campaign Name': campaign_name,
            'Ad Group Name': '',
            'Start Date': current_date,
            'End Date': '',
            'Targeting Type': 'MANUAL',
            'State': 'enabled',
            'Daily Budget': daily_budget,  # Set the daily budget
            'SKU': '',
            'Ad Group Default Bid': '',
            'Bid': '',
            'Keyword Text': '',
            'Match Type': '',
            'Bidding Strategy': 'Dynamic bids - down only'
        }
        final_output_df = append_row(final_output_df, campaign_row)
        
        # Append Ad Group row
        ad_group_row = {
            'Product': 'Sponsored Products',
            'Entity': 'Ad Group',
            'Operation': 'Create',
            'Campaign ID': campaign_name,
            'Ad Group ID': 'EXACT',  # Set to EXACT
            'Portfolio ID': '',
            'Campaign Name': campaign_name,
            'Ad Group Name': 'EXACT',  # Set to EXACT
            'Start Date': '',
            'End Date': '',
            'Targeting Type': '',
            'State': 'enabled',
            'Daily Budget': '',
            'SKU': '',
            'Ad Group Default Bid': '2.00',  # Assuming default bid value is set to 2.00
            'Bid': '',
            'Keyword Text': '',
            'Match Type': '',
            'Bidding Strategy': ''
        }
        final_output_df = append_row(final_output_df, ad_group_row)

        # Append SKU row (Product Ad)
        product_ad_row = {
            'Product': 'Sponsored Products',
            'Entity': 'Product Ad',
            'Operation': 'Create',
            'Campaign ID': campaign_name,
            'Ad Group ID': 'EXACT',  # Set to EXACT
            'Portfolio ID': '',
            'Campaign Name': campaign_name,
            'Ad Group Name': 'EXACT',  # Set to EXACT
            'Start Date': '',
            'End Date': '',
            'Targeting Type': '',
            'State': 'enabled',
            'Daily Budget': '',
            'SKU': sku,
            'Ad Group Default Bid': '',
            'Bid': '',
            'Keyword Text': '',
            'Match Type': '',
            'Bidding Strategy': ''
        }
        final_output_df = append_row(final_output_df, product_ad_row)

        # Append Keyword rows
        added_keywords = set()  # Track added keywords to avoid duplicates
        for keyword in keywords:
            if keyword not in added_keywords:
                avg_cpc_series = avg_cpc_df.loc[(avg_cpc_df['Keyword'] == keyword) & (avg_cpc_df['Portfolio ID'] == portfolio_id), 'Avg CPC']
                if not avg_cpc_series.empty:
                    bid = round(avg_cpc_series.values[0], 2)
                else:
                    bid = overall_avg_cpc  # Set bid to overall average CPC if specific CPC not found
                
                keyword_row = {
                    'Product': 'Sponsored Products',
                    'Entity': 'Keyword',
                    'Operation': 'Create',
                    'Campaign ID': campaign_name,
                    'Ad Group ID': 'EXACT',  # Set to EXACT
                    'Portfolio ID': '',
                    'Campaign Name': campaign_name,
                    'Ad Group Name': 'EXACT',  # Set to EXACT
                    'Start Date': '',
                    'End Date': '',
                    'Targeting Type': '',
                    'State': 'enabled',
                    'Daily Budget': '',
                    'SKU': '',
                    'Ad Group Default Bid': '',
                    'Bid': bid,  # Set the bid
                    'Keyword Text': keyword,
                    'Match Type': 'Exact',
                    'Bidding Strategy': ''
                }
                final_output_df = append_row(final_output_df, keyword_row)
                added_keywords.add(keyword)  # Add to the set

    # Display the final_output_df to verify
    st.write("Final Output DataFrame with 'Entity' column filled:")
    st.dataframe(final_output_df)

    # Provide download link for the final output
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        final_output_df.to_excel(writer, index=False, sheet_name='Final Output')
        writer.close()  # Corrected method to save the file
        st.download_button(label="Download Output Excel", data=buffer.getvalue(), file_name="final_output.xlsx", mime="application/vnd.ms-excel")
