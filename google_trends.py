import pandas as pd
from pytrends.request import TrendReq
import time
import matplotlib.pyplot as plt
import os
from typing import List, Optional
from pathlib import Path
import io
import random  # Am Anfang der Datei hinzufügen

# Configuration
MAX_RETRIES = 3  # Reduziert von 5 auf 3
MIN_RETRY_DELAY = 60  # Minimale Wartezeit in Sekunden
MAX_RETRY_DELAY = 70  # Maximale Wartezeit in Sekunden
BATCH_DELAY = 30  # Wartezeit zwischen Batches
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
    """Fetch Google Trends data with retry logic and batch processing."""
    batch_size = 5
    all_data = []
    
    for i in range(0, len(keywords), batch_size):
        batch_keywords = keywords[i:i + batch_size]
        print(f"Fetching data for keywords: {batch_keywords}")
        
        # Füge zufällige Wartezeit zwischen Batches ein
        if i > 0:
            delay = random.uniform(BATCH_DELAY, BATCH_DELAY + 10)
            print(f"Waiting {delay:.1f} seconds before next batch...")
            time.sleep(delay)
        
        success = False
        for attempt in range(MAX_RETRIES):
            try:
                pytrends.build_payload(batch_keywords, cat=0, timeframe=TIMEFRAME, geo='', gprop='')
                batch_data = pytrends.interest_over_time()
                
                if batch_data is not None and not batch_data.empty:
                    all_data.append(batch_data)
                    print(f"Successfully fetched data for batch {i//batch_size + 1}")
                    success = True
                    break
                
            except Exception as e:
                remaining_attempts = MAX_RETRIES - attempt - 1
                print(f"Error occurred (attempts remaining: {remaining_attempts}):")
                print(e)
                
                if remaining_attempts > 0:
                    # Zufällige Wartezeit zwischen MIN_RETRY_DELAY und MAX_RETRY_DELAY
                    retry_delay = random.uniform(MIN_RETRY_DELAY, MAX_RETRY_DELAY)
                    print(f"Retrying in {retry_delay:.1f} seconds...")
                    time.sleep(retry_delay)
                else:
                    print(f"Failed to fetch data for batch: {batch_keywords}")
        
        if not success:
            print(f"Skipping batch after {MAX_RETRIES} failed attempts")
            continue
    
    if not all_data:
        return None
    
    # Combine all batches
    combined_data = pd.concat(all_data, axis=1)
    # Remove duplicate columns if any
    combined_data = combined_data.loc[:, ~combined_data.columns.duplicated()]
    return combined_data

def save_data(data: pd.DataFrame, keywords: List[str]) -> None:
    """Save data to CSV and create/save plot."""
    # Create output directory if it doesn't exist
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Save CSV - reset_index() to include the date column properly
    csv_path = os.path.join(OUTPUT_DIR, 'google_trends_data.csv')
    data.reset_index().to_csv(csv_path, index=False)
    print(f"Data saved to '{csv_path}'")
    
    # Create and save plot
    plt.figure(figsize=(10, 6))
    # Plot directly from the original DataFrame
    for keyword in keywords:
        if keyword in data.columns:  # Check if keyword exists in the data
            plt.plot(data.index, data[keyword], label=keyword)
    
    plt.title("Google Trends Data")
    plt.xlabel("Date")
    plt.ylabel("Search Interest")
    plt.legend(title="Keywords")
    plt.grid()
    
    plot_path = os.path.join(OUTPUT_DIR, 'google_trends_plot.png')
    plt.savefig(plot_path)
    plt.close()  # Close the figure to free memory
    print(f"Plot saved to '{plot_path}'")

def main():
    """Main execution function."""
    try:
        # Initialize pytrends with a custom timeout
        pytrends = TrendReq(hl='en-US', tz=360, timeout=(10, 25))  # Connect timeout: 10s, Read timeout: 25s
        
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