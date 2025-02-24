import streamlit as st
# Must be the first Streamlit command
st.set_page_config(page_title="Google Trends Analyzer", layout="wide")

try:
    import pandas as pd
    from google_trends import (
        load_keywords_from_file,
        fetch_trends_data,
        TrendReq,
        OUTPUT_DIR
    )
    import matplotlib.pyplot as plt
    import os
    from datetime import datetime
    from PIL import Image
    import plotly.express as px
    from concurrent.futures import ThreadPoolExecutor
    import numpy as np
    
    st.success("All imports successful!")
except Exception as e:
    st.error(f"Import error: {str(e)}")
    st.stop()

st.title("Google Trends Analyzer")

st.write("""
Upload an Excel file containing keywords to analyze Google Trends data.
The Excel file should have a column named 'Keywords'.
""")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

@st.cache_data
def load_keywords_cached(file):
    """Cache the keyword loading process"""
    return load_keywords_from_file(file)

@st.cache_data(ttl=3600)
def cached_fetch_trends_data(keywords_batch):
    """Cached version of fetch_trends_data for a batch of keywords"""
    try:
        return fetch_trends_data(keywords_batch)
    except Exception as e:
        st.error(f"Error in cached_fetch_trends_data: {str(e)}")
        return None

def process_keyword_batch(keywords_batch, status_text):
    """Process a batch of keywords together"""
    try:
        status_text.text(f"Fetching data for keywords: {', '.join(keywords_batch)}")
        return cached_fetch_trends_data(keywords_batch)
    except Exception as e:
        st.error(f"Error in process_keyword_batch: {str(e)}")
        return None

if uploaded_file is not None:
    try:
        # Load keywords with caching
        keywords = load_keywords_cached(uploaded_file)
        
        st.write("Keywords found:", keywords)
        
        # Add a button to start analysis
        if st.button("Analyze Trends"):
            # Create a progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Process all keywords at once
                status_text.text("Fetching trends data...")
                data = fetch_trends_data(keywords)
                progress_bar.progress(1.0)
                
                if data is not None and not data.empty:
                    # Create two columns for CSV and plot
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write("### Trend Data")
                        st.dataframe(data)
                        
                        # Create download button for CSV
                        csv = data.to_csv(index=True)
                        st.download_button(
                            label="Download CSV",
                            data=csv,
                            file_name="google_trends_data.csv",
                            mime="text/csv"
                        )
                    
                    with col2:
                        st.write("### Trend Visualization")
                        # Convert data to long format for plotly
                        data_long = data.reset_index().melt(
                            id_vars=['date'],
                            value_vars=list(data.columns),
                            var_name='Keyword',
                            value_name='Interest'
                        )
                        
                        # Create interactive plot with plotly
                        fig = px.line(
                            data_long,
                            x='date',
                            y='Interest',
                            color='Keyword',
                            title="Google Trends Data"
                        )
                        fig.update_layout(
                            xaxis_title="Date",
                            yaxis_title="Search Interest",
                            hovermode='x unified'
                        )
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # Create download button for plot
                        fig.write_image("temp_plot.png")
                        with open('temp_plot.png', 'rb') as file:
                            st.download_button(
                                label="Download Plot",
                                data=file,
                                file_name="google_trends_plot.png",
                                mime="image/png"
                            )
                        # Clean up temporary file
                        os.remove('temp_plot.png')
                else:
                    st.error("No data retrieved. Please check your keywords or try again later.")
            
            except Exception as e:
                st.error(f"Error during data fetching: {str(e)}")
            
            finally:
                # Clear the progress indicators
                status_text.empty()
                progress_bar.empty()
    
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

# Add logo to sidebar first
try:
    logo = Image.open('kovacs.png')
    st.sidebar.image(logo, width=200)
except Exception as e:
    st.sidebar.error(f"Could not load logo: {str(e)}")

# Then show instructions
st.sidebar.write("""
### Instructions
1. Prepare an Excel file with a column named 'Keywords'
2. Upload the file using the file uploader
3. Click 'Analyze Trends' to fetch and visualize the data
4. Download the results as CSV or PNG
""")