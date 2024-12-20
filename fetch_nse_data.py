import yfinance as yf
import pandas as pd

# Define the list of ticker symbols for multiple NSE stocks
ticker_symbols = ["RELIANCE.NS", "TCS.NS", "INFY.NS", "HDFCBANK.NS", "ICICIBANK.NS"]

# Initialize an empty DataFrame to store close prices
close_prices = pd.DataFrame()

# Fetch data for each ticker symbol and store the close prices in the DataFrame
for ticker in ticker_symbols:
    try:
        data = yf.download(ticker, period="1y")
        close_prices[ticker] = data['Close']
    except Exception as e:
        print(f"Unable to fetch data for {ticker}: {e}")
        # If unable to fetch data, fill the column with 0
        close_prices[ticker] = 0

# Save the close prices to an Excel file
close_prices.to_excel("NSE_stocks_close_prices.xlsx")

print("Close prices for multiple NSE stocks for the last 1 year have been downloaded and saved to NSE_stocks_close_prices.xlsx")
