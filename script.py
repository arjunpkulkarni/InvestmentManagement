import pandas as pd
import yfinance as yf

# Fetch financial data for a single ticker
def fetch_direct_financial_data(ticker):
    try:
        stock = yf.Ticker(ticker)
        info = stock.info

        financial_data = {
            "Ticker": ticker,
            "P/E Ratio": info.get("trailingPE"),
            "P/B Ratio": info.get("priceToBook"),
            "Dividend Yield (%)": info.get("dividendYield", 0) * 100 if info.get("dividendYield") else None,
            "Market Cap": info.get("marketCap"),
        }
        return financial_data
    except Exception as e:
        print(f"Error fetching data for {ticker}: {e}")
        return None

# Save results to an Excel file
def process_tickers_to_excel(tickers, output_file):
    results = []

    for ticker in tickers:
        print(f"Processing {ticker}...")
        data = fetch_direct_financial_data(ticker)
        if data:
            results.append(data)

    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

# Example tickers
tickers = ["TSLA", "AAPL", "MSFT", "AMZN", "GOOGL"]
output_file = "Financial_Analysis.xlsx"
process_tickers_to_excel(tickers, output_file)
