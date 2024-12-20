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
failed_tickers = []

def fetch_data(ticker):
    try:
        data = yf.download(ticker, start=start_date, end=end_date)[['Close']]
        if not data.empty:
            close_prices[ticker] = data['Close']
        else:
            close_prices[ticker] = pd.Series(0, index=pd.date_range(start=start_date, end=end_date))
            failed_tickers.append(ticker)
    except Exception as e:
        close_prices[ticker] = pd.Series(0, index=pd.date_range(start=start_date, end=end_date))
        failed_tickers.append(ticker)
        print(f"Failed to fetch data for {ticker}: {e}")

# Use ThreadPoolExecutor to fetch data concurrently
with ThreadPoolExecutor(max_workers=10) as executor:
    executor.map(fetch_data, tickers)

# Combine all close prices into a DataFrame
close_prices_df = pd.DataFrame(close_prices)

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter('nse_750_stocks_close_last_one_year.xlsx', engine='xlsxwriter') as writer:
    close_prices_df.to_excel(writer, sheet_name='Close Prices')

# Print the details of failed tickers
if failed_tickers:
    print("Failed to fetch data for the following tickers (considered value as 0):")
    for ticker in failed_tickers:
        print(ticker)
else:
    print("Successfully fetched data for all tickers.")

print("Historical data for 750 NSE stocks (Close prices) for the last one year has been generated and saved to 'nse_750_stocks_close_last_one_year.xlsx'.")
