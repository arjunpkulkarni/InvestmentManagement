import pandas as pd
import yfinance as yf

# Define thresholds for metrics
THRESHOLDS = {
    "P/E Ratio": 15,          # Price-to-Earnings
    "P/B Ratio": 3.0,         # Price-to-Book
    "PEG Ratio": 1.0,         # Price/Earnings-to-Growth
    "Dividend Yield (%)": 4.0, # Minimum Dividend Yield
    "Earnings Yield (%)": 5.0, # Minimum Earnings Yield
    "EV/EBITDA": 13.0,        # Enterprise Value to EBITDA
    "Free Cash Flow (%)": 4.0, # Free Cash Flow per share/Price
    "Interest Coverage": 1.5, # Minimum Interest Coverage Ratio
    "Operating Margin (%)": 12.0, # Operating Profit Margin
}

# Fetch financial data for a given ticker
def fetch_financial_data(ticker):
    try:
        stock = yf.Ticker(ticker)
        info = stock.info

        # Extract financial data
        financial_data = {
            "Ticker": ticker,
            "P/E Ratio": info.get("trailingPE"),
            "P/B Ratio": info.get("priceToBook"),
            "PEG Ratio": info.get("pegRatio"),
            "Dividend Yield (%)": info.get("dividendYield", 0) * 100 if info.get("dividendYield") else None,
            "Earnings Yield (%)": 100 / info.get("trailingPE") if info.get("trailingPE") else None,
            "EV/EBITDA": info.get("enterpriseToEbitda"),
            "Free Cash Flow (%)": None,  # Placeholder, requires calculation
            "Interest Coverage": None,  # Placeholder, requires calculation
            "Operating Margin (%)": info.get("operatingMargins", 0) * 100 if info.get("operatingMargins") else None,
        }

        # Calculate Free Cash Flow and Interest Coverage
        cashflow = info.get("freeCashflow")
        market_cap = info.get("marketCap")
        if cashflow and market_cap:
            financial_data["Free Cash Flow (%)"] = (cashflow / market_cap) * 100

        ebit = info.get("ebitda")
        interest_expense = info.get("interestExpense")
        if ebit and interest_expense:
            financial_data["Interest Coverage"] = ebit / interest_expense

        # Evaluate investment decision
        financial_data["Decision"] = evaluate_investment(financial_data)

        return financial_data
    except Exception as e:
        print(f"Error fetching data for {ticker}: {e}")
        return None

# Evaluate investment based on thresholds
def evaluate_investment(data):
    # Check all thresholds
    if (
        data["P/E Ratio"] and data["P/E Ratio"] <= THRESHOLDS["P/E Ratio"]
        and data["P/B Ratio"] and data["P/B Ratio"] <= THRESHOLDS["P/B Ratio"]
        and data["PEG Ratio"] and data["PEG Ratio"] <= THRESHOLDS["PEG Ratio"]
        and data["Dividend Yield (%)"] and data["Dividend Yield (%)"] >= THRESHOLDS["Dividend Yield (%)"]
        and data["Earnings Yield (%)"] and data["Earnings Yield (%)"] >= THRESHOLDS["Earnings Yield (%)"]
        and data["EV/EBITDA"] and data["EV/EBITDA"] <= THRESHOLDS["EV/EBITDA"]
        and data["Free Cash Flow (%)"] and data["Free Cash Flow (%)"] >= THRESHOLDS["Free Cash Flow (%)"]
        and data["Interest Coverage"] and data["Interest Coverage"] >= THRESHOLDS["Interest Coverage"]
        and data["Operating Margin (%)"] and data["Operating Margin (%)"] >= THRESHOLDS["Operating Margin (%)"]
    ):
        return "Invest"
    return "Hold"

# Process multiple tickers and export results
def process_tickers_to_excel(tickers, output_file):
    results = []

    for ticker in tickers:
        print(f"Processing {ticker}...")
        data = fetch_financial_data(ticker)
        if data:
            results.append(data)

    # Create DataFrame and export to Excel
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

# List of tickers (replace with your own tickers)
tickers = ['TSLA', 'MSFT', 'GOOGL', 'MRNA', 'PFE', 'AAPL', 'BA', 'O', 'WMT', 'COIN',
 'NEE', 'NVDA', 'AMZN', 'LMT', 'VRTX', 'SONY', 'GS', 'SHOP', 'RIVN', 'ENPH',
 'SPCE', 'MARA', 'ISRG', 'META', 'SQ', 'PLD', 'ETSY', 'NIO', 'XOM', 'CRM',
 'BABA', 'FSLR', 'BITF', 'RKLB', 'PYPL', 'SPG', 'CHPT', 'RIOT', 'VICI', 'LCID', 'VZ', 'TXN', 'BILI', 'QCOM', 'SE', 'RIO', 'OKTA', 'BAH', 'PANW', 'SPOT',
 'INTC', 'IBM', 'TMUS', 'DOCU', 'HUBS', 'T', 'BIDU', 'NFLX', 'SHOP', 'NTES',
 'WIX', 'JD', 'MU', 'UBER', 'ZM', 'AMD', 'LYFT', 'TTWO', 'EBAY', 'TME',
 'PINS', 'SQSP', 'ASML', 'SNOW', 'RBLX', 'AVGO', 'NVCR', 'EA', 'ADBE', 'DDOG',
 'WDC', 'CRWD', 'ZS', 'CRSP', 'CMCSA', 'BHP', 'BEAM', 'ORCL', 'TWLO', 'CSCO']



# Output file name
output_file = "Comprehensive_Stock_Analysis.xlsx"

# Run the script
process_tickers_to_excel(tickers, output_file)
