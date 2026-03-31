import yfinance as yf
import pandas as pd
import os

def pull_data(ticker_symbol="AAPL"):
    ticker = yf.Ticker(ticker_symbol)

    income_stmt = ticker.financials        # income statement
    balance_sheet = ticker.balance_sheet   # balance sheet
    cash_flow = ticker.cashflow            # cash flow statement

    
    os.makedirs("data", exist_ok=True)

    income_stmt.T.to_csv("data/income_statement.csv")
    balance_sheet.T.to_csv("data/balance_sheet.csv")
    cash_flow.T.to_csv("data/cash_flow.csv")

    print(f"Data pulled for {ticker_symbol} and saved to /data")

if __name__ == "__main__":
    pull_data()