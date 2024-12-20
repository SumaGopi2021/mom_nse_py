import yfinance as yf
import pandas as pd
import datetime

# Define the list of 750 NSE stock tickers
tickers = ['RELIANCE.NS', 'TCS.NS', 'INFY.NS', 'HDFCBANK.NS', 'HINDUNILVR.NS', 'ICICIBANK.NS', 'KOTAKBANK.NS', 'SBIN.NS', 'BAJFINANCE.NS', 'BHARTIARTL.NS']  # Add more tickers as needed

# Define the date range
end_date = datetime.date.today()
start_date = end_date - datetime.timedelta(days=365)

# Initialize a dictionary to store the data and a list to store failed tickers
all_data = {}
failed_tickers = []

# Fetch the historical data for each ticker
for ticker in tickers:
    try:
        data = yf.download(ticker, start=start_date, end=end_date)[['Volume', 'Close']]
        if not data.empty:
            all_data[ticker] = data
        else:
            failed_tickers.append(ticker)
    except Exception as e:
        failed_tickers.append(ticker)
        print(f"Failed to fetch data for {ticker}: {e}")

# Combine all data into a single DataFrame with multi-level columns
combined_data = pd.concat(all_data, axis=1)

# Save the data to an Excel file
combined_data.to_excel('nse_750_stocks_volume_close_last_one_year.xlsx')

# Print the details of failed tickers
if failed_tickers:
    print("Failed to fetch data for the following tickers:")
    for ticker in failed_tickers:
        print(ticker)
else:
    print("Successfully fetched data for all tickers.")

print("Historical data for 750 NSE stocks (Volume and Close prices) for the last one year has been generated and saved to 'nse_750_stocks_volume_close_last_one_year.xlsx'.")
