# streamlit_ngram_analysis.py

import streamlit as st

try:
    import pandas as pd
    import nltk
    from nltk.corpus import stopwords
    from nltk.util import ngrams
    from collections import Counter
    from io import BytesIO
    from scipy.stats import pearsonr
except ModuleNotFoundError as e:
    st.error(f"An error occurred while importing modules: {e}")

# Ensure nltk resources are downloaded
nltk.download('punkt')
nltk.download('stopwords')

# Get the list of stop words
stop_words = set(stopwords.words('english'))

# Function to preprocess the text
def preprocess_text(text):
    tokens = nltk.word_tokenize(text.lower())
    # Remove stop words
    filtered_tokens = [word for word in tokens if word.isalnum() and word not in stop_words]
    return filtered_tokens

# Function to generate n-grams
def generate_ngrams(text, n):
    tokens = preprocess_text(text)
    return list(ngrams(tokens, n))

# Function to analyze n-grams
def analyze_ngrams(data, column, n):
    all_ngrams = []
    for text in data[column].dropna():
        all_ngrams.extend(generate_ngrams(text, n))
    
    ngram_counts = Counter(all_ngrams)
    ngram_df = pd.DataFrame(ngram_counts.items(), columns=['N-gram', 'Count'])
    ngram_df = ngram_df.sort_values(by='Count', ascending=False)
    return ngram_df

# Function to convert DataFrame to Excel for download
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='N-gram Analysis')
    processed_data = output.getvalue()
    return processed_data

# Function to identify high-performing keywords
def identify_high_performing_keywords(data, ngram_df, performance_column):
    high_performing_ngrams = ngram_df.head(10)  # Top 10 n-grams
    high_performing_keywords = data[data[performance_column].isin([' '.join(ngram) for ngram in high_performing_ngrams['N-gram']])]
    return high_performing_keywords

# Function to calculate correlation between two metrics for n-grams
def calculate_correlation(data, ngram_df, text_column, metric1, metric2):
    correlations = []
    for ngram, count in ngram_df[['N-gram', 'Count']].values:
        ngram_str = ' '.join(ngram)
        filtered_data = data[data[text_column].str.contains(ngram_str, case=False, na=False)]
        if len(filtered_data) > 1:  # Ensure there are at least two data points
            correlation, _ = pearsonr(filtered_data[metric1], filtered_data[metric2])
            correlations.append((ngram_str, correlation))
        else:
            correlations.append((ngram_str, None))  # Append None if not enough data points
    correlation_df = pd.DataFrame(correlations, columns=['N-gram', f'Correlation {metric1} & {metric2}'])
    return correlation_df

# Streamlit app
st.title("N-gram Analysis on Search Terms Performance Report")

# File uploader
uploaded_file = st.file_uploader("Upload Amazon Bulk File", type="xlsx")

if uploaded_file is not None:
    try:
        # Read the uploaded file
        data = pd.read_excel(uploaded_file)
        
        # Display the first few rows of the uploaded file
        st.write("Uploaded Data Preview")
        st.dataframe(data.head())
        
        # Select the text column for analysis
        text_column = st.selectbox("Select the column for N-gram analysis", data.columns)
        
        # Select the first performance column for analysis
        metric1 = st.selectbox("Select the first performance metric", data.columns)
        
        # Select the second performance column for correlation analysis
        metric2 = st.selectbox("Select the second performance metric", data.columns)
        
        # Select the n-gram length
        n = st.slider("Select N-gram length", 1, 5, 2)
        
        # Perform the n-gram analysis
        ngram_df = analyze_ngrams(data, text_column, n)
        
        # Display the first few rows of the n-gram analysis
        st.write(f"N-gram Analysis (N={n})")
        st.dataframe(ngram_df.head(20))
        
        # Identify high-performing keywords
        high_performing_keywords = identify_high_performing_keywords(data, ngram_df, metric1)
        
        # Display the high-performing keywords
        st.write("High-Performing Keywords based on N-gram Analysis")
        st.dataframe(high_performing_keywords.head(20))
        
        # Calculate and display correlation between the two selected metrics
        correlation_df = calculate_correlation(data, ngram_df, text_column, metric1, metric2)
        
        # Display the correlation analysis
        st.write(f"Correlation between {metric1} and {metric2}")
        st.dataframe(correlation_df.head(20))
        
        # Provide download link for the n-gram analysis and correlation file
        st.write("Download the N-gram analysis and correlation file:")
        combined_df = pd.concat([ngram_df, correlation_df], axis=1)
        st.download_button(
            label="Download Excel file",
            data=to_excel(combined_df),
            file_name="ngram_analysis_with_correlation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Insights
        st.write("Insights:")
        most_common_ngram = ngram_df.iloc[0]['N-gram']
        st.write(f"The most common {n}-gram is: {' '.join(most_common_ngram)} with {ngram_df.iloc[0]['Count']} occurrences.")
    
    except Exception as e:
        st.error(f"An error occurred: {e}")

else:
    st.info("Please upload an Amazon Bulk File to start the analysis.")
