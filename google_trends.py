import pandas as pd
from pytrends.request import TrendReq
import time
import matplotlib.pyplot as plt
import os
from typing import List, Optional
import random
import requests

# Print version information
print("Python version:", pd.__version__)
print("Pytrends version:", TrendReq.__module__)

# Configuration
OUTPUT_DIR = 'output'
EXCEL_FILE = 'keywords.xlsx'
SHEET_NAME = 'Sheet1'

# Make sure these are available for import
__all__ = ['TrendReq', 'OUTPUT_DIR', 'fetch_trends_data', 'load_keywords_from_file']

def load_keywords_from_file(file_path: str) -> List[str]:
    """Load keywords from Excel file."""
    df = pd.read_excel(file_path)
    return df['Keywords'].dropna().tolist()

def test_internet_connection():
    """Test if we can connect to Google."""
    try:
        response = requests.get('https://trends.google.com', timeout=5)
        print(f"Connection test status code: {response.status_code}")
        return response.status_code == 200
    except Exception as e:
        print(f"Connection test failed: {str(e)}")
        return False

def fetch_trends_data(keywords: List[str]) -> Optional[pd.DataFrame]:
    """Simple function to get trends data."""
    print("\nDEBUG: fetch_trends_data called with keywords:", keywords)
    print("DEBUG: Type of keywords:", type(keywords))
    
    if not isinstance(keywords, list):
        print("ERROR: keywords must be a list")
        return None
        
    if not keywords:
        print("ERROR: keywords list is empty")
        return None
        
    if not all(isinstance(k, str) for k in keywords):
        print("ERROR: all keywords must be strings")
        return None

    try:
        print("\nInitializing TrendReq...")
        pytrends = TrendReq(
            hl='de-DE',
            tz=60,
            timeout=(10, 25),
            retries=2,
            backoff_factor=0.5
        )
        
        print(f"\nBuilding payload for keywords: {keywords}")
        pytrends.build_payload(
            kw_list=keywords[:1],  # Just take the first keyword
            timeframe='today 12-m',
            geo='DE'
        )
        
        print("\nRequesting data from Google Trends...")
        data = pytrends.interest_over_time()
        
        if data is not None and not data.empty:
            print("SUCCESS: Data retrieved!")
            print(f"Data shape: {data.shape}")
            print(f"Columns: {data.columns.tolist()}")
            return data
        else:
            print("ERROR: No data returned from Google Trends")
            return None
            
    except Exception as e:
        print(f"ERROR in fetch_trends_data: {type(e).__name__}: {str(e)}")
        return None

def analyze_trends(keywords: List[str]) -> Optional[pd.DataFrame]:
    """Simple function to get trends data."""
    print(f"Starting trends analysis for keywords: {keywords}")
    
    try:
        # Initialize pytrends
        pytrends = TrendReq(hl='en-US', tz=360)
        
        # Build payload
        pytrends.build_payload(
            keywords,
            cat=0,
            timeframe='today 5-y',
            geo='DE',
            gprop=''
        )
        
        # Get data
        data = pytrends.interest_over_time()
        
        if data is not None and not data.empty:
            print("Successfully retrieved data!")
            return data
        else:
            print("No data returned from Google Trends")
            return None
            
    except Exception as e:
        print(f"Error getting trends data: {str(e)}")
        return None

def save_results(data: pd.DataFrame, keywords: List[str]):
    """Save results to CSV and create plot."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Save CSV
    csv_path = os.path.join(OUTPUT_DIR, 'trends_data.csv')
    data.to_csv(csv_path)
    print(f"Data saved to {csv_path}")
    
    # Create plot
    plt.figure(figsize=(12, 6))
    for kw in keywords:
        if kw in data.columns:
            plt.plot(data.index, data[kw], label=kw)
    
    plt.title('Google Trends Data')
    plt.legend()
    plt.grid(True)
    
    # Save plot
    plot_path = os.path.join(OUTPUT_DIR, 'trends_plot.png')
    plt.savefig(plot_path)
    plt.close()
    print(f"Plot saved to {plot_path}")

def main():
    """Main function."""
    try:
        # Load keywords
        keywords = load_keywords_from_file(EXCEL_FILE)
        print(f"Loaded keywords: {keywords}")
        
        # Get data
        data = fetch_trends_data(keywords)
        
        if data is not None and not data.empty:
            save_results(data, keywords)
            print("\nAnalysis completed successfully!")
        else:
            print("\nCould not complete analysis - no data retrieved")
            
    except Exception as e:
        print(f"Error in main function: {str(e)}")

if __name__ == "__main__":
    main()