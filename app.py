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

# Add this near the top of the file, after imports
@st.cache_data(ttl=3600)  # Cache data for 1 hour
def cached_fetch_trends_data(keyword):
    """Cached version of fetch_trends_data for single keyword"""
    pytrends = TrendReq(hl='en-US', tz=360)
    return fetch_trends_data(pytrends, [keyword])

if uploaded_file is not None:
    try:
        # Load keywords
        keywords = load_keywords_from_file(uploaded_file)
        
        st.write("Keywords found:", keywords)
        
        # Add a button to start analysis
        if st.button("Analyze Trends"):
            # Create a progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Calculate progress steps
            total_keywords = len(keywords)
            progress_step = 1.0 / total_keywords
            
            # Initialize empty DataFrame for data
            data = pd.DataFrame()
            
            # Process each keyword with progress updates
            for idx, keyword in enumerate(keywords):
                status_text.text(f"Fetching data for keyword: {keyword} ({idx+1}/{total_keywords})")
                try:
                    # Use cached function instead of direct API call
                    keyword_data = cached_fetch_trends_data(keyword)
                    if keyword_data is not None and not keyword_data.empty:
                        if data.empty:
                            data = keyword_data
                        else:
                            data = pd.merge(data, keyword_data, left_index=True, right_index=True, how='outer')
                    
                    # Update progress
                    progress_bar.progress((idx + 1) * progress_step)
                    
                except Exception as e:
                    st.warning(f"Error fetching data for {keyword}: {str(e)}")
                    continue
            
            # Clear the status text and progress bar
            status_text.empty()
            progress_bar.empty()
            
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
                    fig, ax = plt.subplots(figsize=(10, 6))
                    data[keywords].plot(ax=ax, title="Google Trends Data")
                    plt.xlabel("Date")
                    plt.ylabel("Search Interest")
                    plt.legend(title="Keywords")
                    plt.grid()
                    st.pyplot(fig)
                    
                    # Create download button for plot
                    plt.savefig('temp_plot.png')
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