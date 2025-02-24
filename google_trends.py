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
    """Function to get trends data with improved error handling."""
    print(f"Starting trends analysis for keywords: {keywords}")
    
    try:
        # Use provided pytrends instance or create new one
        if pytrends is None:
            pytrends = TrendReq(hl='en-US', tz=360, timeout=(10, 25), retries=2)
        
        # Process keywords in smaller batches (2 at a time)
        batch_size = 2
        batches = [keywords[i:i + batch_size] for i in range(0, len(keywords), batch_size)]
        all_data = []
        
        for batch in batches:
            print(f"Processing batch: {batch}")
            try:
                # Add a small delay between batches
                time.sleep(random.uniform(1, 2))
                
                pytrends.build_payload(
                    batch,
                    cat=0,
                    timeframe='today 5-y',  # 5 years of data
                    geo='DE',  # Germany
                    gprop=''
                )
                
                data = pytrends.interest_over_time()
                
                if data is not None and not data.empty:
                    print(f"✓ Successfully fetched data for: {batch}")
                    all_data.append(data)
                else:
                    print(f"No data returned for batch: {batch}")
            
            except Exception as e:
                print(f"Error processing batch {batch}: {str(e)}")
                continue
        
        if not all_data:
            print("No data could be retrieved for any keywords")
            return None
        
        # Combine all batch results
        combined_data = pd.concat(all_data, axis=1)
        # Remove duplicate columns if any
        combined_data = combined_data.loc[:, ~combined_data.columns.duplicated()]
        
        print(f"Successfully retrieved data for {combined_data.shape[1]} keywords")
        return combined_data
            
    except Exception as e:
        print(f"Error in fetch_trends_data: {str(e)}")
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
        if kw in data.columns:  # Only plot if we have data for this keyword
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
        print("Loading keywords...")
        keywords = load_keywords_from_file(EXCEL_FILE)
        print(f"Loaded keywords: {keywords}")
        
        # Get data
        print("\nFetching trends data...")
        data = fetch_trends_data(keywords)
        
        # Save results if we got data
        if data is not None and not data.empty:
            save_results(data, keywords)
            print("\nAnalysis completed successfully!")
        else:
            print("\nCould not complete analysis - no data retrieved")
            
    except Exception as e:
        print(f"Error in main function: {str(e)}")

if __name__ == "__main__":
    main()