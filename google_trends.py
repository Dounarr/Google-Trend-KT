import pandas as pd
from pytrends.request import TrendReq
import matplotlib.pyplot as plt
import os
from typing import List, Optional
import time
import random

# Configuration
OUTPUT_DIR = 'output'
EXCEL_FILE = 'keywords.xlsx'

def load_keywords_from_file(file_path: str) -> List[str]:
    """Load keywords from Excel file."""
    df = pd.read_excel(file_path)
    return df['Keywords'].dropna().tolist()

def fetch_trends_data(keywords: List[str]) -> Optional[pd.DataFrame]:
    """Get trends data with shorter timeframe."""
    print(f"\nStarting trends analysis for: {keywords}")
    
    try:
        # Initialize with minimal settings
        pytrends = TrendReq(
            hl='de-DE',
            tz=60
        )
        
        # Try different timeframes
        timeframes = [
            'today 3-m',  # Last 3 months
            'today 1-m',  # Last month
            'now 7-d'     # Last week
        ]
        
        for timeframe in timeframes:
            try:
                print(f"\nTrying timeframe: {timeframe}")
                
                # Build payload with single keyword
                print("Building payload...")
                pytrends.build_payload(
                    kw_list=keywords[:1],  # Just try first keyword
                    timeframe=timeframe,
                    geo='DE'
                )
                
                # Get the data
                print("Requesting data...")
                data = pytrends.interest_over_time()
                
                if data is not None and not data.empty:
                    print("Successfully retrieved data!")
                    print(f"Data shape: {data.shape}")
                    print(f"Columns: {data.columns.tolist()}")
                    return data
                else:
                    print(f"No data for timeframe {timeframe}")
                    
                # Add delay between attempts
                delay = random.uniform(5, 10)
                print(f"Waiting {delay:.1f} seconds...")
                time.sleep(delay)
                
            except Exception as e:
                print(f"Error with timeframe {timeframe}: {str(e)}")
                continue
        
        print("\nAll timeframes failed")
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