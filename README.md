# Python-Automated Financial Model — Apple Inc. (AAPL)

A Python script that automatically pulls Apple's financial data, builds a 
3-statement financial model, runs a DCF valuation, and outputs everything 
into a formatted Excel workbook — with one command.

## What it does

- Pulls 4 years of historical financials from Yahoo Finance using `yfinance`
- Builds an Income Statement, Balance Sheet, and Cash Flow Statement
- Calculates key metrics: Gross Margin, EBIT Margin, Net Margin, Free Cash Flow
- Runs a DCF valuation with projected FCFs, WACC, and terminal value
- Outputs a formatted, multi-sheet Excel workbook with charts and a summary page

## Output

The generated `financial_model.xlsx` contains 6 sheets:
| Sheet | Contents |
|---|---|
| Summary | Key metrics snapshot — FY2025 vs FY2024 + DCF intrinsic value |
| Income Statement | Revenue, Gross Profit, EBIT, EBITDA, Net Income + margins |
| Balance Sheet | Assets, Liabilities, Equity, Debt, Cash |
| Cash Flow | Operating CF, Capex, FCF, D&A |
| DCF Valuation | 5-year FCF projections + discounted values |
| Chart | Revenue vs Free Cash Flow bar chart |

## Tech stack

| Tool | Purpose |
|---|---|
| `yfinance` | Pull financial statements from Yahoo Finance |
| `pandas` | Clean and reshape financial data |
| `openpyxl` | Write and format Excel workbook |
| `matplotlib` | Generate embedded charts |

## How to run

**1. Clone the repo**
```bash
git clone https://github.com/veralaiys/automated-financial-model.git
cd automated-financial-model
```

**2. Install dependencies**
```bash
pip install yfinance pandas openpyxl matplotlib
```

**3. Run**
```bash
python data_pull.py   # pulls data from Yahoo Finance → saves to /data
python output.py      # builds model and generates financial_model.xlsx
```

## DCF assumptions

| Assumption | Value |
|---|---|
| WACC | 9% |
| Terminal growth rate | 3% |
| Forecast period | 5 years |
| FCF growth rate | Historical average, capped at 15%, floored at 5% |

These are starting assumptions. Adjust them in `model.py` to run 
different scenarios.

## Project structure
```
automated-financial-model/
│
├── data/                      # pulled financial data (auto-generated)
│   ├── income_statement.csv
│   ├── balance_sheet.csv
│   └── cash_flow.csv
│
├── notebooks/
│   └── exploration.ipynb      # working notebook used during development
│
├── data_pull.py               # pulls and saves raw data
├── model.py                   # financial calculations and DCF logic
├── output.py                  # builds and formats Excel workbook
└── financial_model.xlsx       # final output (auto-generated)
```

## Key concepts demonstrated

- 3-statement financial modelling (IS → BS → CFS linkages)
- DCF valuation (FCF projections, WACC, terminal value, present value)
- Python automation of Excel with `openpyxl`
- Financial data pipeline with `yfinance` and `pandas`
