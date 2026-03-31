import pandas as pd

def load_data():
    income_stmt = pd.read_csv("Data/income_statement.csv", index_col=0)
    balance_sheet = pd.read_csv("data/balance_sheet.csv", index_col=0)
    cash_flow = pd.read_csv("data/cash_flow.csv", index_col=0)

    income_stmt = income_stmt.apply(pd.to_numeric, errors='coerce')
    balance_sheet = balance_sheet.apply(pd.to_numeric, errors='coerce')
    cash_flow = cash_flow.apply(pd.to_numeric, errors='coerce')

    return income_stmt, balance_sheet, cash_flow


def calc_income_metrics(df):
    metrics = pd.DataFrame(index=df.index)
    metrics['Revenue'] = df['Total Revenue']
    metrics['Gross Profit'] = df['Gross Profit']
    metrics['EBIT'] = df['EBIT']
    metrics['EBITDA'] = df['EBITDA']
    metrics['Net Income'] = df['Net Income']
    
    metrics['Gross Margin %'] = df['Gross Profit'] / df['Total Revenue'] * 100
    metrics['EBIT Margin %'] = df['EBIT'] / df['Total Revenue'] * 100
    metrics['Net Margin %'] = df['Net Income'] / df['Total Revenue'] * 100

    metrics = metrics.dropna(how='all')
    return metrics.round(2)


def calc_balance_metrics(df):
    metrics = pd.DataFrame(index=df.index)
    
    metrics['Total Assets'] = df['Total Assets']
    metrics['Total Liabilities'] = df['Total Liabilities Net Minority Interest']
    metrics['Total Equity'] = df['Stockholders Equity']
    metrics['Total Debt'] = df['Total Debt']
    metrics['Cash'] = df['Cash And Cash Equivalents']
    metrics['Net Debt'] = df['Net Debt']

    # check: shd be close to 0 if balance sheet balance_sheet
    metrics['Check(A-L-E)'] = (df['Total Assets'] 
                               - df['Total Liabilities Net Minority Interest'] 
                               - df['Stockholders Equity'])
    
    metrics = metrics.dropna(how='all')
    return metrics.round(2)


def calc_cashflow_metrics(df):
    metrics = pd.DataFrame(index=df.index)

    metrics['Operating CF'] = df['Operating Cash Flow']
    metrics['Capex'] = df['Capital Expenditure']
    metrics['Free Cash Flow'] = df['Free Cash Flow']
    metrics['D&A'] = df['Depreciation And Amortization']

    metrics = metrics.dropna(how='all')
    return metrics.round(2)


def calc_dcf(cf_df, wacc=0.09, terminal_growth=0.03, forecast_years=5):
    fcf_base = cf_df['Free Cash Flow'].iloc[0]
    fcf_series = cf_df['Free Cash Flow'].dropna()
    fcf_growth = fcf_series.pct_change(-1).mean()
    fcf_growth = max(fcf_growth, 0.05)

    projected_fcfs = []
    for year in range(1, forecast_years+1):
        projected_fcf = fcf_base * (1 + fcf_growth) ** year
        discounted_fcf = projected_fcf / (1 + wacc) ** year
        projected_fcfs.append({
            'Year': f'Year {year}',
            'Projected FCF': round(projected_fcf/1e9, 2),
            'Discounted FCF': round(discounted_fcf/1e9, 2)
        })
    
    dcf_df = pd.DataFrame(projected_fcfs)

    fcf_final = fcf_base * (1 + fcf_growth) ** forecast_years
    terminal_value = fcf_final * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_terminal = terminal_value / (1 + wacc) ** forecast_years
    pv_fcfs = sum([row['Discounted FCF'] for row in projected_fcfs])
    intrinsic_value = pv_fcfs + (pv_terminal / 1e9)

    print(f"DCF Intrinsic Value: ${intrinsic_value:.2f}B")
    return dcf_df, intrinsic_value



if __name__ == "__main__":
    income_stmt, balance_sheet, cash_flow = load_data()
    income_metrics = calc_income_metrics(income_stmt)
    balance_metrics = calc_balance_metrics(balance_sheet)
    cf_metrics = calc_cashflow_metrics(cash_flow)
    dcf_table, intrinsic_value = calc_dcf(cf_metrics)
    print("Model calculations complete!")