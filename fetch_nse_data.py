import yfinance as yf
import pandas as pd
import datetime
from concurrent.futures import ThreadPoolExecutor

# Load tickers from the input file
with open('tickers.txt', 'r') as file:
    tickers = [line.strip() for line in file.readlines()]

# Define the date range
end_date = datetime.date.today()
start_date = end_date - datetime.timedelta(days=365)

# Initialize dictionaries to store the data and a list to store failed tickers
close_prices = {}
volumes = {}
failed_tickers = []

def fetch_data(ticker):
    try:
        data = yf.download(ticker, start=start_date, end=end_date)[['Volume', 'Close']]
        if not data.empty:
            close_prices[ticker] = data['Close']
            volumes[ticker] = data['Volume']
        else:
            failed_tickers.append(ticker)
    except Exception as e:
        failed_tickers.append(ticker)
        print(f"Failed to fetch data for {ticker}: {e}")

# Use ThreadPoolExecutor to fetch data concurrently
with ThreadPoolExecutor(max_workers=10) as executor:
    executor.map(fetch_data, tickers)

# Combine all close prices and volumes into separate DataFrames
if close_prices:
    close_prices_df = pd.DataFrame(close_prices)
else:
    close_prices_df = pd.DataFrame(index=pd.date_range(start=start_date, end=end_date))

if volumes:
    volumes_df = pd.DataFrame(volumes)
else:
    volumes_df = pd.DataFrame(index=pd.date_range(start=start_date, end=end_date))

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter('nse_750_stocks_volume_close_last_one_year.xlsx', engine='xlsxwriter') as writer:
    if not close_prices_df.empty:
        close_prices_df.to_excel(writer, sheet_name='Close Prices')
    if not volumes_df.empty:
        volumes_df.to_excel(writer, sheet_name='Volumes')

# Print the details of failed tickers
if failed_tickers:
    print("Failed to fetch data for the following tickers:")
    for ticker in failed_tickers:
        print(ticker)
else:
    print("Successfully fetched data for all tickers.")

print("Historical data for 750 NSE stocks (Close prices and Volumes) for the last one year has been generated and saved to 'nse_750_stocks_volume_close_last_one_year.xlsx'.")
