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
    print(f"\nStarting trends analysis for {len(keywords)} keywords")
    
    try:
        # Initialize with minimal settings
        pytrends = TrendReq(
            hl='de-DE',
            tz=60
        )
        
        # Process keywords in batches of 5
        all_data = []
        for i in range(0, len(keywords), 5):
            keyword_batch = keywords[i:i+5]
            print(f"\nProcessing batch: {keyword_batch}")
            
            # Try different timeframes
            timeframes = [
                'today 3-m',  # Last 3 months
                'today 1-m',  # Last month
                'now 7-d'     # Last week
            ]
            
            batch_success = False
            for timeframe in timeframes:
                try:
                    print(f"\nTrying timeframe: {timeframe}")
                    
                    # Build payload with current batch
                    print("Building payload...")
                    pytrends.build_payload(
                        kw_list=keyword_batch,
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
                        all_data.append(data)
                        batch_success = True
                        break  # Success! Move to next batch
                    else:
                        print(f"No data for timeframe {timeframe}")
                    
                except Exception as e:
                    print(f"Error with timeframe {timeframe}: {str(e)}")
                    continue
                
                finally:
                    # Add delay between attempts
                    delay = random.uniform(5, 10)
                    print(f"Waiting {delay:.1f} seconds...")
                    time.sleep(delay)
            
            if not batch_success:
                print(f"Warning: Could not get data for batch: {keyword_batch}")
            
            # Add delay between batches
            delay = random.uniform(10, 15)
            print(f"Waiting {delay:.1f} seconds between batches...")
            time.sleep(delay)
        
        # Combine all data if we have any
        if all_data:
            # First, ensure all DataFrames have the same index
            common_index = all_data[0].index
            for df in all_data[1:]:
                common_index = common_index.union(df.index)
            
            # Reindex all DataFrames to have the same index
            aligned_data = [df.reindex(common_index) for df in all_data]
            
            # Combine all data
            combined_data = pd.concat(aligned_data, axis=1)
            
            # Remove duplicate columns if any
            combined_data = combined_data.loc[:, ~combined_data.columns.duplicated()]
            
            # Remove the isPartial column if it exists
            if 'isPartial' in combined_data.columns:
                combined_data = combined_data.drop('isPartial', axis=1)
                
            return combined_data
        
        print("\nNo data retrieved for any keyword batch")
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