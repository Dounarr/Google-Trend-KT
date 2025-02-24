import pandas as pd
from pytrends.request import TrendReq
import time
import matplotlib.pyplot as plt
import os
from typing import List, Optional
import random

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

def fetch_trends_data(keywords: List[str], pytrends: Optional[TrendReq] = None) -> Optional[pd.DataFrame]:
    """Simple, direct approach to fetch trends data."""
    print(f"Attempting to fetch data for keywords: {keywords}")
    
    try:
        # Create new pytrends instance with conservative settings
        pytrends = TrendReq(
            hl='de-DE',
            tz=60,
            timeout=(30, 30),
            retries=2,
            backoff_factor=1
        )
        
        # Try with just the first two keywords as a test
        test_keywords = keywords[:2]
        print(f"Testing with keywords: {test_keywords}")
        
        # Build payload with minimal parameters
        pytrends.build_payload(
            test_keywords,
            timeframe='today 12-m',  # Last 12 months instead of 5 years
            geo='DE'
        )
        
        # Get the data
        print("Requesting data from Google Trends...")
        data = pytrends.interest_over_time()
        
        if data is not None and not data.empty:
            print("Successfully retrieved data!")
            print(f"Data shape: {data.shape}")
            print(f"Columns: {data.columns.tolist()}")
            return data
        else:
            print("No data was returned from Google Trends")
            return None
            
    except Exception as e:
        print(f"Error occurred: {str(e)}")
        print("Type of error:", type(e).__name__)
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
        # Test with hardcoded keywords first
        test_keywords = ["Auto", "Berlin"]
        print("\nTesting with simple keywords first...")
        test_data = fetch_trends_data(test_keywords)
        
        if test_data is not None:
            print("\nTest successful! Now trying with actual keywords...")
            
            # Now try with actual keywords
            keywords = load_keywords_from_file(EXCEL_FILE)
            print(f"Loaded keywords: {keywords}")
            
            data = fetch_trends_data(keywords)
            
            if data is not None and not data.empty:
                save_results(data, keywords)
                print("\nAnalysis completed successfully!")
            else:
                print("\nCould not complete analysis - no data retrieved")
        else:
            print("\nTest failed - could not even get test data")
            
    except Exception as e:
        print(f"Error in main function: {str(e)}")

if __name__ == "__main__":
    main()