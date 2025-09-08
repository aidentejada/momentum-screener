######################################################
import numpy as np
import pandas as pd
import math
from scipy.stats import percentileofscore as score
import yfinance as yf
from statistics import mean
######################################################

# define mega-cap stocks that get liquidity premium in position sizing
MEGA_CAPS = ['AAPL', 'MSFT', 'GOOG', 'GOOGL', 'NVDA', 'META']

def datafetcher():
    """
    main data processing function that:
    1. fetches stock data and calculates momentum metrics
    2. computes capm expected returns using beta
    3. calculates percentile rankings for momentum periods
    4. generates high quality momentum (hqm) scores
    5. implements portfolio optimization with position sizing
    """
    global hqm_df
    stocks = pd.read_csv("sp500.csv")
    momentum_stocks = []

    # capm parameters - risk free rate and market return assumptions
    risk_free_rate = 0.041  # current 10yr treasury rate
    market_return = 0.095   # long-term s&p 500 expected return

    # loop through each s&p 500 stock to calculate momentum and expected return metrics
    for stock in stocks['Symbol']:
        try:
            # fetch 1 year of historical data for momentum calculations
            ticker = yf.Ticker(stock)
            hist = ticker.history(period='1y')
            if len(hist) == 0:
                continue
                
            current_price = hist['Close'].iloc[-1]
            
            # get beta from yahoo finance, default to 1.0 if missing
            beta = ticker.info.get('beta') or 1.0
            if not isinstance(beta, (int, float)):
                beta = 1.0
                
            # calculate capm expected return: rf + beta * (market_return - rf)
            expected_return = risk_free_rate + beta * (market_return - risk_free_rate)

            # calculate historical prices for momentum periods (1y, 6m, 3m, 1m)
            price_1yr_ago = hist['Close'].iloc[0]
            price_6mo_ago = hist['Close'].iloc[-126] if len(hist) >= 126 else None
            price_3mo_ago = hist['Close'].iloc[-63] if len(hist) >= 63 else None
            price_1mo_ago = hist['Close'].iloc[-21] if len(hist) >= 21 else None

            # calculate momentum returns for each time period
            one_year_return = (current_price - price_1yr_ago) / price_1yr_ago
            six_mo_return = (current_price - price_6mo_ago) / price_6mo_ago if price_6mo_ago else None
            three_mo_return = (current_price - price_3mo_ago) / price_3mo_ago if price_3mo_ago else None
            one_mo_return = (current_price - price_1mo_ago) / price_1mo_ago if price_1mo_ago else None

            # store all calculated metrics for this stock
            momentum_stocks.append([
                stock, current_price, 0,
                one_year_return, 'N/A',
                six_mo_return, 'N/A',
                three_mo_return, 'N/A',
                one_mo_return, 'N/A',
                0, beta, expected_return
            ])
        except:
            continue

    # create dataframe with all momentum and fundamental data
    columns = [
        "Ticker", "Price", "Shares to Buy",
        "1y Price Return", "1y Return Percentile",
        "6mo Price Return", "6mo Return Percentile",
        "3mo Price Return", "3mo Return Percentile",
        "1mo Price Return", "1mo Return Percentile",
        "HQM Score", "Beta", "Expected Return"
    ]
    hqm_df = pd.DataFrame(momentum_stocks, columns=columns)

    # calculate percentile rankings for each momentum period
    # this ranks stocks relative to each other (0-100 percentile)
    time_periods = ['1y','6mo','3mo','1mo']
    for time_period in time_periods:
        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        clean_series = hqm_df[change_col].dropna()
        for row in hqm_df.index:
            val = hqm_df.loc[row, change_col]
            hqm_df.loc[row, percentile_col] = score(clean_series, val) if pd.notna(val) else 0

    # calculate high quality momentum (hqm) score using time-weighted average
    # weights: 3m=30%, 6m=25%, 1y=25%, 1m=20% (emphasizes medium-term momentum)
    for row in hqm_df.index:
        hqm_df.loc[row, 'HQM Score'] = (
            hqm_df.loc[row, '1y Return Percentile'] * 0.25 +
            hqm_df.loc[row, '6mo Return Percentile'] * 0.25 +
            hqm_df.loc[row, '3mo Return Percentile'] * 0.30 +
            hqm_df.loc[row, '1mo Return Percentile'] * 0.20
        )

    # calculate volatility for each stock using standard deviation of momentum returns
    # cap volatility at 13% for risk management purposes
    for row in hqm_df.index:
        returns = [hqm_df.loc[row, p] for p in ['1y Price Return','6mo Price Return','3mo Price Return','1mo Price Return'] if pd.notna(hqm_df.loc[row, p])]
        hqm_df.loc[row, 'Volatility'] = np.std(returns) if len(returns)>1 else 0.1
        if hqm_df.loc[row, 'Volatility'] > 0.13:
            hqm_df.loc[row, 'Volatility'] = 0.13

    # select top 20 stocks by hqm score for portfolio construction
    hqm_df.sort_values("HQM Score", ascending=False, inplace=True)
    hqm_df = hqm_df[:20].reset_index(drop=True)

    # calculate raw position weights using custom formula:
    # weight = (hqm_score * expected_return) / (volatility + 0.05)
    # mega-caps get 15% liquidity premium due to better execution
    weight_raws = []
    for i in range(len(hqm_df)):
        raw = hqm_df.loc[i, 'HQM Score'] * hqm_df.loc[i, 'Expected Return'] / (hqm_df.loc[i, 'Volatility'] + 0.05)
        if hqm_df.loc[i, 'Ticker'] in MEGA_CAPS:
            raw *= 1.15  # liquidity premium for mega-caps
        weight_raws.append(raw)

    # normalize weights to sum to 100% with 6.5% maximum position size limit
    total_raw = sum(weight_raws) or 1.0
    hqm_df['Weight'] = [min(w / total_raw, 0.065) for w in weight_raws]
    hqm_df['Weight'] = hqm_df['Weight'] / hqm_df['Weight'].sum()
    hqm_df['WeightRaw'] = weight_raws

def portfolio():
    """
    portfolio construction function that:
    1. gets user input for portfolio size
    2. calculates share quantities and position values
    3. displays formatted portfolio allocation table
    4. computes portfolio-level risk and return metrics
    """
    global portfoliosz
    portfoliosz = input("Enter the size of your portfolio: ")
    
    # input validation for portfolio size
    try:
        float(portfoliosz)
    except ValueError:
        print('That is not a number! Try again:')
        portfoliosz = input("Enter the size of your portfolio: ")

    # display portfolio allocation table with key metrics
    print(f"\nCAMP-Enhanced Portfolio Allocation")
    print(f"Total Portfolio: ${float(portfoliosz):,.2f}")
    print(f"{'Rank':<4} {'Ticker':<7} {'HQM':<6} {'ExpRet':<7} {'Vol':<6} {'Weight':<7} {'Shares':<7} {'Value'}")
    print("-"*80)

    # calculate share quantities and position values for each stock
    total_allocated = 0
    for i in range(len(hqm_df)):
        weight = hqm_df.loc[i, 'Weight']
        position_value = float(portfoliosz) * weight
        shares_to_buy = math.floor(position_value / hqm_df.loc[i, "Price"])  # floor to avoid fractional shares
        actual_value = shares_to_buy * hqm_df.loc[i, "Price"]
        hqm_df.loc[i, "Shares to Buy"] = shares_to_buy
        hqm_df.loc[i, "Position Value"] = actual_value
        total_allocated += actual_value
        
        # display formatted row for each position
        print(f"{i+1:<4} {hqm_df.loc[i, 'Ticker']:<7} {hqm_df.loc[i, 'HQM Score']:<6.1f} "
              f"{hqm_df.loc[i, 'Expected Return']:<7.2%} {hqm_df.loc[i, 'Volatility']:<6.3f} "
              f"{weight:<7.1%} {shares_to_buy:<7} ${actual_value:>7,.0f}")

    # calculate portfolio summary statistics
    cash_remaining = float(portfoliosz) - total_allocated
    print("-"*80)
    print(f"Total Invested: ${total_allocated:,.2f}")
    print(f"Cash Remaining: ${cash_remaining:,.2f} ({cash_remaining/float(portfoliosz):.1%})")
    
    # portfolio risk score: weighted average volatility (lower is better)
    weighted_volatility = (hqm_df['Volatility'] * hqm_df['Weight']).sum()
    print(f"Portfolio Risk Score: {weighted_volatility:.3f} (lower is better)")
    
    # portfolio expected return: weighted average of individual stock expected returns
    portfolio_expected_return = (hqm_df['Expected Return'] * hqm_df['Weight']).sum()
    print(f"Portfolio Expected Return (CAPM): {portfolio_expected_return:.2%}")

def xlsx_writer():
    """
    excel export function that:
    1. creates professionally formatted spreadsheet
    2. applies custom styling (dark theme with borders)
    3. formats columns appropriately (currency, percentages, integers)
    4. exports all calculated data for further analysis
    """
    writer = pd.ExcelWriter('momentum_strategy.xlsx', engine='xlsxwriter')
    hqm_df.to_excel(writer, sheet_name='Momentum Strategy', index=False)
    
    # define professional styling templates
    background_color = '#0a0a23'  # dark blue background
    font_color = '#ffffff'        # white text
    string_template = writer.book.add_format({'font_color': font_color, 'bg_color': background_color, 'border':1})
    dollar_template = writer.book.add_format({'num_format':'$0.00','font_color':font_color,'bg_color':background_color,'border':1})
    integer_template = writer.book.add_format({'num_format':'0','font_color':font_color,'bg_color':background_color,'border':1})
    decimal_template = writer.book.add_format({'num_format':'0.00','font_color':font_color,'bg_color':background_color,'border':1})
    percent_template = writer.book.add_format({'num_format':'0.0%','font_color':font_color,'bg_color':background_color,'border':1})

    # define column formatting mapping
    columns_format = {
        "A": ["Ticker", string_template],
        "B": ["Price", dollar_template],
        "C": ["Shares to Buy", integer_template],
        "D": ["1y Price Return", percent_template],
        "E": ["1y Return Percentile", decimal_template],
        "F": ["6mo Price Return", percent_template],
        "G": ["6mo Return Percentile", decimal_template],
        "H": ["3mo Price Return", percent_template],
        "I": ["3mo Return Percentile", decimal_template],
        "J": ["1mo Price Return", percent_template],
        "K": ["1mo Return Percentile", decimal_template],
        "L": ["HQM Score", decimal_template],
        "M": ["Volatility", decimal_template],
        "N": ["Weight", percent_template],
        "O": ["Position Value", dollar_template],
        "P": ["Beta", decimal_template],
        "Q": ["Expected Return", percent_template]
    }

    # apply formatting to each column
    for col in columns_format.keys():
        writer.sheets['Momentum Strategy'].set_column(f'{col}:{col}', 25, columns_format[col][1])
        writer.sheets['Momentum Strategy'].write(f'{col}1', columns_format[col][0], columns_format[col][1])
    writer.close()

# execute the quantitative strategy pipeline
# 1. fetch data and calculate metrics
# 2. construct and display portfolio
# 3. export results to excel
datafetcher()
portfolio()
xlsx_writer()