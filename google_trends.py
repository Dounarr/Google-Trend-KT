import pandas as pd
from pytrends.request import TrendReq
import time
import matplotlib.pyplot as plt
import os
from typing import List, Optional
from pathlib import Path
import io

# Configuration
MAX_RETRIES = 5
RETRY_DELAY = 60  # seconds
TIMEFRAME = 'today 12-m'
OUTPUT_DIR = 'output'
EXCEL_FILE = 'keywords.xlsx'
SHEET_NAME = 'Sheet1'

def load_keywords_from_file(file_object, sheet_name: str = 'Sheet1') -> List[str]:
    """Load keywords from an uploaded Excel file object."""
    try:
        keyword_data = pd.read_excel(file_object, sheet_name=sheet_name)
        
        if 'Keywords' not in keyword_data.columns:
            raise ValueError("Excel file must contain a 'Keywords' column")
        
        keywords = keyword_data['Keywords'].dropna().tolist()
        
        if not keywords:
            raise ValueError("No keywords found in the Excel file")
        
        return keywords
    
    except Exception as e:
        print(f"Error loading keywords: {str(e)}")
        raise

def load_keywords(file_path: str, sheet_name: str) -> List[str]:
    """Load keywords from Excel file with error handling."""
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found: {file_path}")
        
        return load_keywords_from_file(file_path, sheet_name)
    
    except Exception as e:
        print(f"Error loading keywords: {str(e)}")
        raise

def fetch_trends_data(pytrends: TrendReq, keywords: List[str]) -> Optional[pd.DataFrame]:
    """Fetch Google Trends data with retry logic."""
    for attempt in range(MAX_RETRIES):
        try:
            pytrends.build_payload(keywords, cat=0, timeframe=TIMEFRAME, geo='', gprop='')
            data = pytrends.interest_over_time()
            return data
        
        except Exception as e:
            remaining_attempts = MAX_RETRIES - attempt - 1
            print(f"Error occurred (attempts remaining: {remaining_attempts}):")
            print(e)
            
            if remaining_attempts > 0:
                print(f"Retrying in {RETRY_DELAY} seconds...")
                time.sleep(RETRY_DELAY)
            else:
                print("Maximum retry attempts reached.")
                return None

def save_data(data: pd.DataFrame, keywords: List[str]) -> None:
    """Save data to CSV and create/save plot."""
    # Create output directory if it doesn't exist
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Save CSV
    csv_path = os.path.join(OUTPUT_DIR, 'google_trends_data.csv')
    data.to_csv(csv_path)
    print(f"Data saved to '{csv_path}'")
    
    # Create and save plot
    plt.figure(figsize=(10, 6))
    data[keywords].plot(title="Google Trends Data")
    plt.xlabel("Date")
    plt.ylabel("Search Interest")
    plt.legend(title="Keywords")
    plt.grid()
    
    plot_path = os.path.join(OUTPUT_DIR, 'google_trends_plot.png')
    plt.savefig(plot_path)
    plt.show()
    print(f"Plot saved to '{plot_path}'")

def main():
    """Main execution function."""
    try:
        # Initialize pytrends
        pytrends = TrendReq(hl='en-US', tz=360)
        
        # Load keywords
        keywords = load_keywords(EXCEL_FILE, SHEET_NAME)
        
        # Fetch data
        data = fetch_trends_data(pytrends, keywords)
        
        if data is not None and not data.empty:
            save_data(data, keywords)
        else:
            print("No data retrieved. Please check your keywords or time range.")
    
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()