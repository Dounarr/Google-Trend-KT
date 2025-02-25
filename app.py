import streamlit as st
# Must be the first Streamlit command
st.set_page_config(page_title="Google Trends Analyzer", layout="wide")

try:
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
    """Get trends data for all keywords."""
    print(f"\nStarting trends analysis for: {keywords}")
    
    try:
        # Initialize with minimal settings
        pytrends = TrendReq(
            hl='de-DE',
            tz=60
        )
        
        all_data = []
        timeframe = 'today 3-m'  # Use 3-month timeframe
        
        # Process keywords in batches of 5 (Google Trends limit)
        for i in range(0, len(keywords), 5):
            batch = keywords[i:i+5]
            print(f"\nProcessing batch: {batch}")
            
            try:
                # Build payload with batch of keywords
                print("Building payload...")
                pytrends.build_payload(
                    kw_list=batch,
                    timeframe=timeframe,
                    geo='DE'
                )
                
                # Get the data
                print("Requesting data...")
                data = pytrends.interest_over_time()
                
                if data is not None and not data.empty:
                    # Drop isPartial column if it exists
                    if 'isPartial' in data.columns:
                        data = data.drop('isPartial', axis=1)
                    all_data.append(data)
                    print(f"Successfully retrieved data for batch!")
                else:
                    print(f"No data for batch {batch}")
                
                # Add delay between batches
                if i + 5 < len(keywords):  # If not the last batch
                    delay = random.uniform(5, 10)
                    print(f"Waiting {delay:.1f} seconds before next batch...")
                    time.sleep(delay)
                
            except Exception as e:
                print(f"Error processing batch {batch}: {str(e)}")
                continue
        
        if all_data:
            # Combine all batches
            final_data = pd.concat(all_data, axis=1)
            print("\nSuccessfully combined all data!")
            print(f"Final data shape: {final_data.shape}")
            print(f"Columns: {final_data.columns.tolist()}")
            return final_data
        else:
            print("\nNo data was retrieved")
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
    plt