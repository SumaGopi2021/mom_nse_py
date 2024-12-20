import yfinance as yf
import pandas as pd
import datetime

# Load tickers from the input file
with open('tickers.txt', 'r') as file:
    tickers = [line.strip() for line in file.readlines()]

# Define the date range
end_date = datetime.date.today()
start_date = end_date - datetime.timedelta(days=365)

# Initialize dictionaries to store the data and a list to store failed tickers
close_prices = {}
volumes = {}
combined_data = {}
failed_tickers = []

# Fetch the historical data for each ticker
for ticker in tickers:
    try:
        data = yf.download(ticker, start=start_date, end=end_date)[['Volume', 'Close']]
        if not data.empty:
            close_prices[ticker] = data['Close']
            volumes[ticker] = data['Volume']
            combined_data[f'{ticker} Close'] = data['Close']
            combined_data[f'{ticker} Volume'] = data['Volume']
        else:
            failed_tickers.append(ticker)
    except Exception as e:
        failed_tickers.append(ticker)
        print(f"Failed to fetch data for {ticker}: {e}")

# Combine all close prices and volumes into separate DataFrames
close_prices_df = pd.DataFrame(close_prices)
volumes_df = pd.DataFrame(volumes)
combined_df = pd.DataFrame(combined_data)

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter('nse_750_stocks_volume_close_last_one_year.xlsx', engine='xlsxwriter') as writer:
    close_prices_df.to_excel(writer, sheet_name='Close Prices')
    volumes_df.to_excel(writer, sheet_name='Volumes')
    combined_df.to_excel(writer, sheet_name='Combined')

# Print the details of failed tickers
if failed_tickers:
    print("Failed to fetch data for the following tickers:")
    for ticker in failed_tickers:
        print(ticker)
else:
    print("Successfully fetched data for all tickers.")

print("Historical data for 750 NSE stocks (Volume and Close prices) for the last one year has been generated and saved to 'nse_750_stocks_volume_close_last_one_year.xlsx'.")
