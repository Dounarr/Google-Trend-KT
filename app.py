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

if uploaded_file is not None:
    try:
        # Load keywords
        keywords = load_keywords_from_file(uploaded_file)
        
        st.write("Keywords found:", keywords)
        
        # Add a button to start analysis
        if st.button("Analyze Trends"):
            with st.spinner('Fetching Google Trends data...'):
                # Initialize pytrends
                pytrends = TrendReq(hl='en-US', tz=360)
                
                # Fetch data
                data = fetch_trends_data(pytrends, keywords)
                
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