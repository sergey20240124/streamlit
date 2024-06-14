#streamlit run ta_analysis_new.py
import streamlit as st
import pandas as pd
import warnings
from datetime import timedelta
from io import BytesIO

# Suppress the specific UserWarning from openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Data cleaning function with correct calculations for CTR and ACOS
def clean_amazon_data(file):
    try:
        amazon_data = pd.read_excel(file)
        amazon_data.columns = amazon_data.columns.str.strip()
        amazon_data['Date'] = pd.to_datetime(amazon_data['Date'])
        amazon_data.fillna(0, inplace=True)

        columns_to_remove = [
            'Top-of-search Impression Share', 
            'Total Return on Advertising Spend (ROAS)', '7 Day Total Orders (#)',
            '7 Day Conversion Rate', '7 Day Advertised SKU Units (#)',
            '7 Day Other SKU Units (#)', '7 Day Advertised SKU Sales', '7 Day Other SKU Sales',
            'Currency', 'Ad Group Name', 'Total Advertising Cost of Sales (ACOS)', 'Click-Thru Rate (CTR)'
        ]
        amazon_data.drop(columns=columns_to_remove, inplace=True)

        amazon_data.rename(columns={
            '7 Day Total Sales': 'Ad Sales', 
            'Spend': 'Ad Spend', 
            '7 Day Total Units (#)': 'Units',
            'Cost Per Click (CPC)': 'CPC',
            'Portfolio name': 'Portfolio'
        }, inplace=True)

        numeric_columns = ['Impressions', 'Ad Sales', 'Units', 'Clicks', 'Ad Spend', 'CPC']
        amazon_data[numeric_columns] = amazon_data[numeric_columns].apply(pd.to_numeric, errors='coerce').round(2)

        # Calculate CTR and ACOS as numeric values
        amazon_data['CTR'] = (amazon_data['Clicks'] / amazon_data['Impressions']).round(4)
        amazon_data['ACOS'] = (amazon_data['Ad Spend'] / amazon_data['Ad Sales']).round(2)

        amazon_data = amazon_data[amazon_data['Ad Spend'] > 0]

        exclude_targeting = ['loose-match', 'close-match', 'complements', 'substitutes']
        amazon_data = amazon_data[~amazon_data['Targeting'].isin(exclude_targeting)]

        exclude_keywords = ['category', 'B0']
        pattern = '|'.join(exclude_keywords)
        amazon_data = amazon_data[~amazon_data['Targeting'].str.contains(pattern, case=False, na=False)]

        amazon_data.reset_index(drop=True, inplace=True)
        
        return amazon_data

    except Exception as e:
        st.error(f"Error in cleaning data: {e}")
        return pd.DataFrame()

# Function to create the 'Review' sheet
def create_review_sheet(cleaned_data):
    try:
        end_date = cleaned_data['Date'].max()
        start_date = end_date - timedelta(weeks=6)
        recent_data = cleaned_data[(cleaned_data['Date'] > start_date) & (cleaned_data['Date'] <= end_date)]
        
        ad_spend = recent_data.groupby(['Portfolio', pd.Grouper(key='Date', freq='W-SUN')])['Ad Spend'].sum().unstack(fill_value=0)
        ad_sales = recent_data.groupby(['Portfolio', pd.Grouper(key='Date', freq='W-SUN')])['Ad Sales'].sum().unstack(fill_value=0)
        acos = (ad_spend / ad_sales).replace([float('inf'), -float('inf'), pd.NA], 0).round(2)
        
        review_data = pd.concat([ad_spend, acos], axis=1, keys=['Spend', 'ACOS'])
        review_data.columns = [f"{col[1].strftime('%m-%d-%Y')} {col[0]}" for col in review_data.columns]
        
        columns = ['Portfolio']
        dates = sorted(set(col.split()[0] for col in review_data.columns))

        # Add Spend columns first
        for date in dates:
            columns.append(f"{date} Spend")
        
        # Add ACOS columns after Spend columns
        for date in dates:
            columns.append(f"{date} ACOS")
        
        review_data = review_data.reset_index()
        review_data = review_data[columns]
        
        last_week_col = [col for col in review_data.columns if 'Spend' in col][-1]
        review_data.sort_values(by=last_week_col, ascending=False, inplace=True)
        
        return review_data

    except Exception as e:
        st.error(f"Error in creating review sheet: {e}")
        return pd.DataFrame()

# Function to create individual portfolio sheets with Ad Sales, Ad Spend, and ACOS columns
def create_portfolio_sheets(cleaned_data):
    try:
        portfolio_sheets = {}

        end_date = cleaned_data['Date'].max()
        start_date = end_date - timedelta(weeks=6)
        recent_data = cleaned_data[(cleaned_data['Date'] > start_date) & (cleaned_data['Date'] <= end_date)]

        portfolios = recent_data['Portfolio'].unique()

        for portfolio in portfolios:
            try:
                portfolio_data = recent_data[recent_data['Portfolio'] == portfolio]

                portfolio_sheet = portfolio_data.groupby(['Targeting', pd.Grouper(key='Date', freq='W-SUN')]).agg(
                    Ad_Sales=('Ad Sales', 'sum'),
                    Ad_Spend=('Ad Spend', 'sum'),
                    ACOS=('ACOS', 'mean')  # Assuming you want to average the ACOS over the week
                ).unstack(fill_value=0)

                portfolio_sheet.columns = ['_'.join([col[0], col[1].strftime('%m-%d-%Y')]) for col in portfolio_sheet.columns]
                portfolio_sheet.reset_index(inplace=True)

                # Collect all the unique dates present in the columns
                dates = sorted(set(col.split('_')[1] for col in portfolio_sheet.columns if '_' in col))

                for date in dates:
                    spend_col = f'Ad_Spend_{date}'
                    sales_col = f'Ad_Sales_{date}'
                    acos_col = f'ACOS_{date}'

                    # Add the Spend and ACOS columns if they don't already exist
                    if spend_col in portfolio_sheet.columns:
                        portfolio_sheet[f'Spend_{date}'] = portfolio_sheet[spend_col]
                    if acos_col in portfolio_sheet.columns:
                        portfolio_sheet[f'ACOS_{date}'] = portfolio_sheet[acos_col]

                columns_to_keep = ['Targeting'] + [col for col in portfolio_sheet.columns if 'Ad_Sales' in col or 'Spend' in col or 'ACOS' in col]
                portfolio_sheet = portfolio_sheet[columns_to_keep]

                portfolio_sheets[portfolio] = portfolio_sheet

            except Exception as e:
                st.error(f"Error processing portfolio {portfolio}: {e}")

        return portfolio_sheets

    except Exception as e:
        st.error(f"Error in creating portfolio sheets: {e}")
        return {}

# Function to create the 'Spend Tracking' sheet with adjusted calculations
def create_spend_tracking_sheet(cleaned_data):
    try:
        end_date = cleaned_data['Date'].max()
        start_date = end_date - timedelta(weeks=5)
        last_week_start = end_date - timedelta(weeks=1)
        last_week_end = end_date

        # Data for the prior 4 weeks excluding the last week
        prior_4_weeks_data = cleaned_data[(cleaned_data['Date'] > start_date) & (cleaned_data['Date'] <= last_week_start)].copy()
        # Data for the last week
        last_week_data = cleaned_data[(cleaned_data['Date'] > last_week_start) & (cleaned_data['Date'] <= last_week_end)].copy()

        # Add a new column to indicate the week
        prior_4_weeks_data.loc[:, 'Week'] = prior_4_weeks_data['Date'].dt.isocalendar().week
        last_week_data.loc[:, 'Week'] = last_week_data['Date'].dt.isocalendar().week

        # Group by Portfolio, Campaign Name, Targeting, and Week to calculate total spend, sales, and CPC
        spend_by_keyword = prior_4_weeks_data.groupby(['Portfolio', 'Campaign Name', 'Targeting', 'Week']).agg(
            weekly_spend=('Ad Spend', 'sum'),
            weekly_sales=('Ad Sales', 'sum'),
            weekly_cpc=('CPC', 'mean')
        ).reset_index()

        # Calculate the number of weeks active and the total spend per keyword
        spend_summary = spend_by_keyword.groupby(['Portfolio', 'Campaign Name', 'Targeting']).agg(
            total_spend=('weekly_spend', 'sum'),
            total_sales=('weekly_sales', 'sum'),
            weeks_active=('Week', 'count'),
            avg_cpc=('weekly_cpc', 'mean')  # Average CPC for the prior 4 weeks
        ).reset_index()

        # Calculate the average spend and average sales by dividing the total by the number of weeks active
        spend_summary['4 Week Avg Spend'] = spend_summary['total_spend'] / spend_summary['weeks_active']
        spend_summary['4 Week Avg Sales'] = spend_summary['total_sales'] / spend_summary['weeks_active']

        # Calculate the average ACOS as 4 Week Avg Spend / 4 Week Avg Sales
        spend_summary['4 Week Avg ACOS'] = (spend_summary['4 Week Avg Spend'] / spend_summary['4 Week Avg Sales']).replace([float('inf'), -float('inf')], 0).round(2)

        # Calculate the last week's spend, sales, and CPC
        last_week_spend = last_week_data.groupby(['Portfolio', 'Campaign Name', 'Targeting']).agg(
            last_week_spend=('Ad Spend', 'sum'),
            last_week_sales=('Ad Sales', 'sum'),
            last_week_cpc=('CPC', 'mean')
        ).reset_index()

        # Calculate the last week's ACOS as Last Week Spend / Last Week Sales
        last_week_spend['last_week_acos'] = (last_week_spend['last_week_spend'] / last_week_spend['last_week_sales']).replace([float('inf'), -float('inf')], 0).round(2)

        # Merge the average spend and last week's spend into a single DataFrame
        spend_tracking = spend_summary.merge(
            last_week_spend.rename(columns={'last_week_spend': 'Last Week Spend', 'last_week_sales': 'Last Week Sales', 'last_week_cpc': 'Last Week CPC', 'last_week_acos': 'Last Week ACOS'}),
            on=['Portfolio', 'Campaign Name', 'Targeting'],
            how='left'
        ).fillna(0)

        # Calculate the change percentage
        spend_tracking['Change'] = ((spend_tracking['Last Week Spend'] - spend_tracking['4 Week Avg Spend']) / spend_tracking['4 Week Avg Spend']).replace([float('inf'), -float('inf')], 0).round(2)

        # Determine the status based on the change percentage
        spend_tracking['Status'] = spend_tracking['Change'].apply(
            lambda x: 'Increase 25% plus' if x > 0.25 else ('Decrease 25% plus' if x < -0.25 else 'Stable')
        )

        # Filter out 'Stable' keywords and those with zero spend last week
        spend_tracking = spend_tracking[(spend_tracking['Status'] != 'Stable') & (spend_tracking['Last Week Spend'] > 0)]

        # Select and order the final columns
        spend_tracking = spend_tracking[['Portfolio', 'Campaign Name', 'Targeting', 'Status', '4 Week Avg Spend', 'Last Week Spend', 
                                         '4 Week Avg Sales', 'Last Week Sales', '4 Week Avg ACOS', 'Last Week ACOS', 'avg_cpc', 'Last Week CPC']]
        spend_tracking.rename(columns={'avg_cpc': '4 Week Avg CPC'}, inplace=True)

        # Round all values to 2 decimal places
        spend_tracking = spend_tracking.round(2)

        return spend_tracking

    except Exception as e:
        st.error(f"Error in creating spend tracking sheet: {e}")
        return pd.DataFrame()

# Function to adjust column widths
def adjust_column_widths(writer):
    workbook = writer.book
    for sheetname in workbook.sheetnames:
        worksheet = workbook[sheetname]
        for col in worksheet.columns:
            max_length = max(len(str(cell.value)) for cell in col) + 1
            max_length = min(max_length, 30)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max_length
    
# Function to convert DataFrame to Excel in memory and adjust column widths
def to_excel(cleaned_data, review_data, portfolio_sheets, spend_tracking_data):
    try:
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl')

        spend_tracking_data.to_excel(writer, index=False, sheet_name='Spend_Tracking')
        review_data.to_excel(writer, index=False, sheet_name='Review')
        cleaned_data.to_excel(writer, index=False, sheet_name='Base')

        for portfolio, data in portfolio_sheets.items():
            # Replace invalid characters in sheet names
            safe_portfolio = "".join([c if c.isalnum() or c in [' ', '_'] else "_" for c in portfolio])
            data.to_excel(writer, index=False, sheet_name=safe_portfolio)

        # Adjust column widths
        adjust_column_widths(writer)
        
        writer.close()
        processed_data = output.getvalue()

        return processed_data

    except Exception as e:
        st.error(f"Error in converting to Excel: {e}")
        return None
           

# Streamlit app
def main():
    st.title("Keyword Performance Tracker")
    st.markdown("""
This app generates trackers to analyze targeting performance.<br><br>
Input file - SP Targeting Report<br>
Time Unit - Daily<br><br>
Content<br>
>Spend Tracking: Targets having 25%+ Spend increase/decline<br>
Review: Portfolios overall performance (Spend, ACOS)<br>
Base: Cleaned report data for reference<br>
Separate Portfolio Sheets<br>

""", unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload Amazon SP Targeting Report (Make sure the time unit is Daily)", type="xlsx")

    if uploaded_file is not None:
        st.write("File uploaded successfully!")

        cleaned_data = clean_amazon_data(uploaded_file)
        if cleaned_data.empty:
            st.error("Failed to clean data.")
            return
        
        review_data = create_review_sheet(cleaned_data)
        portfolio_sheets = create_portfolio_sheets(cleaned_data)
        spend_tracking_data = create_spend_tracking_sheet(cleaned_data)

        if review_data.empty:
            st.error("Failed to create review sheet.")
            return

        if spend_tracking_data.empty:
            st.error("Failed to create spend tracking sheet.")
            return

        #st.write("Preview of review data:")
        #st.dataframe(review_data)

        st.write("Preview of spend tracking data:")
        st.dataframe(spend_tracking_data)

        # Option to view individual portfolio sheets
        #st.write("Preview of individual portfolio sheets:")
        #for portfolio, data in portfolio_sheets.items():
            #st.write(f"Portfolio: {portfolio}")
            #st.dataframe(data.head())

        cleaned_data_excel = to_excel(cleaned_data, review_data, portfolio_sheets, spend_tracking_data)

        if cleaned_data_excel:
            st.download_button(
                label="Download XLSX file",
                data=cleaned_data_excel,
                file_name='Target_Review.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

if __name__ == "__main__":
    main()
