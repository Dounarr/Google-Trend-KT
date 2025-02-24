import pandas as pd
from pytrends.request import TrendReq
import time
import matplotlib.pyplot as plt
import os
from typing import List, Optional, Dict
import asyncio
from concurrent.futures import ThreadPoolExecutor
import random
import pickle
from datetime import datetime, timedelta
import hashlib
import threading

# Configuration
MAX_RETRIES = 3
BATCH_DELAY = 5  # Further reduced delay
OUTPUT_DIR = 'output'
EXCEL_FILE = 'keywords.xlsx'
SHEET_NAME = 'Sheet1'
CACHE_DIR = 'cache'
CACHE_DURATION = timedelta(days=1)  # Cache results for 1 day

def get_cache_key(keywords: List[str], timeframe: str, geo: str) -> str:
    """Generate a unique cache key for the query."""
    query_string = f"{'-'.join(sorted(keywords))}-{timeframe}-{geo}"
    return hashlib.md5(query_string.encode()).hexdigest()

def get_cached_data(cache_key: str) -> Optional[pd.DataFrame]:
    """Retrieve cached data if it exists and is not expired."""
    cache_file = os.path.join(CACHE_DIR, f"{cache_key}.pkl")
    if os.path.exists(cache_file):
        modification_time = datetime.fromtimestamp(os.path.getmtime(cache_file))
        if datetime.now() - modification_time < CACHE_DURATION:
            with open(cache_file, 'rb') as f:
                return pickle.load(f)
    return None

def save_to_cache(cache_key: str, data: pd.DataFrame) -> None:
    """Save data to cache."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_file = os.path.join(CACHE_DIR, f"{cache_key}.pkl")
    with open(cache_file, 'wb') as f:
        pickle.dump(data, f)

async def fetch_batch(batch_keywords: List[str], session_pool: Dict[int, TrendReq]) -> Optional[pd.DataFrame]:
    """Fetch a single batch of keywords using session pool."""
    thread_id = threading.get_ident()
    
    # Get or create session for this thread
    if thread_id not in session_pool:
        session_pool[thread_id] = TrendReq(hl='en-US', tz=360, timeout=(10, 25), retries=2)
    
    pytrends = session_pool[thread_id]
    
    # Check cache first
    cache_key = get_cache_key(batch_keywords, 'today 5-y', 'DE')
    cached_data = get_cached_data(cache_key)
    if cached_data is not None:
        print(f"✓ Retrieved from cache: {batch_keywords}")
        return cached_data

    for attempt in range(MAX_RETRIES):
        try:
            pytrends.build_payload(
                batch_keywords,
                cat=0,
                timeframe='today 5-y',
                geo='DE',
                gprop=''
            )
            data = pytrends.interest_over_time()
            
            if data is not None and not data.empty:
                print(f"✓ Successfully fetched data for: {batch_keywords}")
                save_to_cache(cache_key, data)
                return data
            
            print(f"No data returned for: {batch_keywords}")
            
        except Exception as e:
            print(f"Error on attempt {attempt + 1} for {batch_keywords}: {str(e)}")
            if attempt < MAX_RETRIES - 1:
                await asyncio.sleep(random.uniform(1, 3))
    
    return None

async def fetch_trends_data(keywords: List[str]) -> Optional[pd.DataFrame]:
    """Fetch Google Trends data using async approach with session pool."""
    batch_size = 2
    batches = [keywords[i:i + batch_size] for i in range(0, len(keywords), batch_size)]
    all_data = []
    session_pool = {}
    
    print(f"\nProcessing {len(batches)} batches of keywords...")
    
    # Create a thread pool for running pytrends requests
    with ThreadPoolExecutor(max_workers=5) as executor:  # Increased workers
        loop = asyncio.get_event_loop()
        
        # Process batches in chunks to control concurrency
        chunk_size = 3
        for i in range(0, len(batches), chunk_size):
            chunk = batches[i:i + chunk_size]
            
            # Create tasks for this chunk
            tasks = []
            for batch in chunk:
                task = loop.run_in_executor(
                    executor,
                    lambda b=batch: fetch_batch(b, session_pool)
                )
                tasks.append(task)
            
            # Wait for chunk to complete
            chunk_results = await asyncio.gather(*tasks)
            all_data.extend([r for r in chunk_results if r is not None])
            
            # Small delay between chunks
            if i + chunk_size < len(batches):
                await asyncio.sleep(random.uniform(1, 2))

    if not all_data:
        print("\nNo data could be retrieved. Possible reasons:")
        print("• Keywords have very low search volume")
        print("• Keywords are too specific")
        print("• Rate limiting from Google Trends")
        return None

    # Combine all results
    combined_data = pd.concat(all_data, axis=1)
    combined_data = combined_data.loc[:, ~combined_data.columns.duplicated()]
    print(f"\nSuccessfully retrieved data for {combined_data.shape[1]} keywords")
    return combined_data

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
        data = fetch_trends_data(keywords)
        
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