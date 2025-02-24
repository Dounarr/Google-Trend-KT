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
        
        # Add validation for empty strings and whitespace
        keywords = [k.strip() for k in keywords if isinstance(k, str) and k.strip()]
        
        if not keywords:
            raise ValueError("No valid keywords found in the Excel file")
        
        print(f"Found {len(keywords)} valid keywords in the Excel file")
        return keywords
    
    except Exception as e:
        print(f"Error loading keywords: {str(e)}")
        raise

def load_keywords(file_path: str, sheet_name: str) -> List[str]:
    """Load keywords from Excel file with error handling."""
    try:
        # Convert to absolute path and check if file exists
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Excel file not found at: {abs_path}")
        
        print(f"Loading keywords from: {abs_path}")
        return load_keywords_from_file(abs_path, sheet_name)
    
    except Exception as e:
        print(f"Error loading keywords: {str(e)}")
        raise

def fetch_trends_data(pytrends: TrendReq, keywords: List[str]) -> Optional[pd.DataFrame]:
    """Fetch Google Trends data with retry logic and batch processing."""
    # Reduce batch size to 2 for better success rate
    batch_size = 2
    all_data = []
    
    for i in range(0, len(keywords), batch_size):
        batch_keywords = keywords[i:i + batch_size]
        print(f"\nFetching data for keywords: {batch_keywords}")
        print(f"Processing batch {i//batch_size + 1} of {(len(keywords) + batch_size - 1)//batch_size}")
        
        # Increase initial delay
        delay = random.uniform(BATCH_DELAY, BATCH_DELAY + 10)
        print(f"Waiting {delay:.1f} seconds before fetching...")
        time.sleep(delay)
        
        success = False
        for attempt in range(MAX_RETRIES):
            try:
                print(f"\nAttempt {attempt + 1}/{MAX_RETRIES}")
                # Try with a longer timeframe to increase chances of getting data
                pytrends.build_payload(
                    batch_keywords,
                    cat=0,
                    timeframe='today 5-y',  # Changed from 12-m to 5-y
                    geo='DE',  # Added German region since these seem to be German keywords
                    gprop=''
                )
                batch_data = pytrends.interest_over_time()
                
                if batch_data is not None and not batch_data.empty:
                    print(f"Success! Data shape: {batch_data.shape}")
                    print(f"Data found for keywords: {batch_data.columns.tolist()}")
                    all_data.append(batch_data)
                    success = True
                    break
                else:
                    print(f"No data returned for keywords: {batch_keywords}")
                    
            except Exception as e:
                print(f"Error occurred: {type(e).__name__}: {str(e)}")
                if attempt < MAX_RETRIES - 1:
                    retry_delay = random.uniform(MIN_RETRY_DELAY, MAX_RETRY_DELAY)
                    print(f"Retrying in {retry_delay:.1f} seconds...")
                    time.sleep(retry_delay)
        
        if not success:
            print(f"Warning: Could not fetch data for keywords: {batch_keywords}")
            continue

    if not all_data:
        print("\nDetailed troubleshooting information:")
        print("- Keywords attempted:", keywords)
        print("- No data was successfully retrieved for any keyword combination")
        print("- This might be because:")
        print("  * The keywords have very low search volume")
        print("  * The keywords are too specific or regional")
        print("  * Google Trends doesn't have enough data for these terms")
        return None
    
    # Combine all batches
    combined_data = pd.concat(all_data, axis=1)
    combined_data = combined_data.loc[:, ~combined_data.columns.duplicated()]
    print(f"\nFinal combined data shape: {combined_data.shape}")
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
        # Print current working directory
        print(f"Current working directory: {os.getcwd()}")
        
        # Check if Excel file exists
        excel_path = os.path.abspath(EXCEL_FILE)
        print(f"Looking for Excel file at: {excel_path}")
        
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found at: {excel_path}")
        
        # Modify the pytrends initialization to increase timeout
        pytrends = TrendReq(hl='en-US', tz=360, timeout=(30, 45), retries=3, backoff_factor=0.5)
        
        # Load keywords
        keywords = load_keywords(EXCEL_FILE, SHEET_NAME)
        print(f"Loaded {len(keywords)} keywords: {keywords}")
        
        # Fetch data
        data = fetch_trends_data(pytrends, keywords)
        
        if data is not None and not data.empty:
            save_data(data, keywords)
        else:
            print("\nNo data retrieved. Please check your keywords or try again later.")
            print("This might be due to:")
            print("1. Rate limiting from Google Trends")
            print("2. Invalid or unsearchable keywords")
            print("3. Network connectivity issues")
            print("\nTroubleshooting steps:")
            print("1. Verify your keywords.xlsx file is properly formatted")
            print("2. Try with fewer keywords (start with 2-3)")
            print("3. Wait a few minutes before trying again")
            print("4. Check your internet connection")
    
    except Exception as e:
        print(f"An error occurred: {type(e).__name__}: {str(e)}")

if __name__ == "__main__":
    main()