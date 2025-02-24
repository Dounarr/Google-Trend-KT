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
    """Function to get trends data with improved error handling and debugging."""
    print(f"Starting trends analysis for keywords: {keywords}")
    
    try:
        # Use provided pytrends instance or create new one with longer timeout
        if pytrends is None:
            pytrends = TrendReq(
                hl='de-DE',  # Changed to German
                tz=60,       # Changed timezone to Europe
                timeout=(30, 30),  # Increased timeout
                retries=3,
                backoff_factor=0.5
            )
        
        # Process keywords in pairs
        batch_size = 2
        batches = [keywords[i:i + batch_size] for i in range(0, len(keywords), batch_size)]
        all_data = []
        
        for batch in batches:
            print(f"\nTrying batch: {batch}")
            
            for attempt in range(3):  # Try each batch up to 3 times
                try:
                    # Add a small delay between attempts
                    if attempt > 0:
                        time.sleep(5)
                    
                    print(f"Building payload for: {batch}")
                    pytrends.build_payload(
                        batch,
                        cat=0,
                        timeframe='today 5-y',
                        geo='DE',
                        gprop=''
                    )
                    
                    print("Requesting data...")
                    data = pytrends.interest_over_time()
                    
                    if data is not None and not data.empty:
                        print(f"✓ Successfully got data for: {batch}")
                        print(f"Data shape: {data.shape}")
                        print(f"Columns: {data.columns.tolist()}")
                        all_data.append(data)
                        break
                    else:
                        print(f"No data returned for batch: {batch}")
                
                except Exception as e:
                    print(f"Attempt {attempt + 1} failed for {batch}: {str(e)}")
                    if attempt == 2:  # Last attempt
                        print(f"All attempts failed for batch: {batch}")
            
            # Add delay between batches
            time.sleep(random.uniform(2, 3))
        
        if not all_data:
            print("\nNo data could be retrieved for any keywords")
            return None
        
        # Combine all batch results
        print("\nCombining results...")
        combined_data = pd.concat(all_data, axis=1)
        
        # Remove duplicate columns if any
        combined_data = combined_data.loc[:, ~combined_data.columns.duplicated()]
        
        print(f"\nFinal data shape: {combined_data.shape}")
        print(f"Final columns: {combined_data.columns.tolist()}")
        
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