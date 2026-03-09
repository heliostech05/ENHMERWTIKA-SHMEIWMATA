import csv
import json
from datetime import datetime, timezone, timedelta
import os
from urllib.parse import urlencode
from urllib.request import urlopen
from urllib.error import URLError, HTTPError

try:
    import requests
except ImportError:
    requests = None

try:
    import pandas as pd
except ImportError:
    pd = None

# Function to get the date range from the start of current year until today
def get_date_range():
    """
    Calculates the date range from the first day of the current year up to today.
    Returns two date strings in ISO 8601 format.
    """
    today = datetime.today()
    start_date = datetime(today.year, 1, 1)  # January 1st of current year
    start_date_str = start_date.strftime('%Y-%m-%d')
    end_date_str = today.strftime('%Y-%m-%d')  # Today
    return start_date_str, end_date_str

# Function to fetch the day-ahead prices from the Energy-Charts API
def fetch_day_ahead_prices(country="GR", start_date="", end_date=""):
    """
    Fetches the day-ahead price data for a given country and date range from the Energy-Charts API.
    
    Parameters:
    - country: The country code (default is "GR" for Greece).
    - start_date: The start date in YYYY-MM-DD format.
    - end_date: The end date in YYYY-MM-DD format.
    
    Returns:
    - A list of timestamps and corresponding prices for the specified country and date range.
    """
    # API endpoint for Day-Ahead Price data
    url = "https://api.energy-charts.info/price"
    
    # Parameters for the API request
    params = {
        "bzn": country,  # Bidding zone (Greece: "GR")
        "start": start_date,  # Start date
        "end": end_date,  # End date
    }
    
    try:
        if requests is not None:
            # Preferred path when requests is installed
            response = requests.get(url, params=params, timeout=60)
            response.raise_for_status()
            data = response.json()
        else:
            # Fallback path: no external dependency
            query = urlencode(params)
            full_url = f"{url}?{query}"
            with urlopen(full_url, timeout=60) as response:
                data = json.loads(response.read().decode("utf-8"))

        # Check if 'price' data is available
        if "price" in data and "unix_seconds" in data:
            return data["unix_seconds"], data["price"]

        print("No data available for the specified range.")
        return [], []
    except (HTTPError, URLError) as e:
        print(f"Error fetching data from the API: {e}")
        return [], []
    except Exception as e:
        print(f"Unexpected error while fetching API data: {e}")
        return [], []

# Function to convert Unix timestamps to human-readable dates
def convert_unix_to_date(unix_seconds):
    """
    Converts Unix timestamps to ISO 8601 dates in UTC+02:00 format.
    
    Parameters:
    - unix_seconds: List of Unix timestamps.
    
    Returns:
    - A list of formatted date strings corresponding to the timestamps.
    """
    utc_plus_2 = timezone(timedelta(hours=2))
    return [
        datetime.fromtimestamp(ts, timezone.utc)
        .astimezone(utc_plus_2)
        .isoformat(timespec='minutes')
        for ts in unix_seconds
    ]

# Function to save the data to a CSV file
def save_data_to_csv(timestamps, prices, filename=None):
    """
    Saves the timestamps and prices data to a CSV file.
    
    Parameters:
    - timestamps: List of human-readable timestamps.
    - prices: List of prices corresponding to each timestamp.
    - filename: The name of the output CSV file.
    """
    if filename is None:
        current_year = datetime.today().year
        filename = f"energy-charts_Electricity_production_and_spot_prices_in_Greece_in_{current_year}.csv"

    # Ensure the output folder exists next to this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(script_dir, "output")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Save the DataFrame to CSV
    output_path = os.path.join(output_folder, filename)
    if pd is not None:
        df = pd.DataFrame({
            'Timestamp': timestamps,
            'Price (EUR/MWh)': prices
        })
        df.to_csv(output_path, index=False)
    else:
        with open(output_path, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Timestamp", "Price (EUR/MWh)"])
            writer.writerows(zip(timestamps, prices))

    print(f"Data successfully saved to {output_path}")

# Main function to fetch and save the day-ahead price data
def main():
    """
    Main function to fetch the day-ahead prices for Greece from the beginning
    of the current year until today,
    and save the data to a CSV file.
    """
    # Step 1: Get the date range from start of current year
    start_date, end_date = get_date_range()
    print(f"Fetching day-ahead price data for Greece from {start_date} to {end_date}...")

    # Step 2: Fetch the price data from the API
    timestamps, prices = fetch_day_ahead_prices(country="GR", start_date=start_date, end_date=end_date)

    # Step 3: If data was fetched successfully, convert timestamps and save the data
    if timestamps and prices:
        # Convert Unix timestamps to human-readable format
        human_readable_timestamps = convert_unix_to_date(timestamps)
        
        # Step 4: Save the data to a CSV file
        save_data_to_csv(human_readable_timestamps, prices)

if __name__ == "__main__":
    # Run the script
    main()
