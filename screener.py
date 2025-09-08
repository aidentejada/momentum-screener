######################################################
import numpy as np                                   #
import pandas as pd                                  #
import math                                          #
from scipy.stats import percentileofscore as score   #
from xlsxwriter import *                             #
import yfinance as yf                                #
from statistics import mean                          #
######################################################

def datafetcher():
    global hqm_df
    # import list of stocks
    stocks = pd.read_csv("sp500.csv")
    momentum_stocks = []

    for stock in stocks['Symbol']:
        try:
            ticker = yf.Ticker(stock)
            hist = ticker.history(period='1y')  # get full year of data
            
            if len(hist) == 0:
                continue
            
            # current price
            current_price = hist['Close'].iloc[-1]
            
            # pulls prices for different periods
            price_1yr_ago = hist['Close'].iloc[0]
            price_6mo_ago = hist['Close'].iloc[-126] if len(hist) >= 126 else None  # ~6 months = 126 trading days
            price_3mo_ago = hist['Close'].iloc[-63] if len(hist) >= 63 else None    # ~3 months = 63 trading days
            price_1mo_ago = hist['Close'].iloc[-21] if len(hist) >= 21 else None    # ~1 month = 21 trading days
            
            # calculates returns
            one_year_return = (current_price - price_1yr_ago) / price_1yr_ago
            six_mo_return = (current_price - price_6mo_ago) / price_6mo_ago if price_6mo_ago else None
            three_mo_return = (current_price - price_3mo_ago) / price_3mo_ago if price_3mo_ago else None
            one_mo_return = (current_price - price_1mo_ago) / price_1mo_ago if price_1mo_ago else None
            
            # building the data frame
            momentum_stocks.append([
                stock, 
                current_price, 
                0,  # shares to buy (for now)
                one_year_return,
                'N/A', # placeholder for percentiles
                six_mo_return,
                'N/A',
                three_mo_return, 
                'N/A',
                one_mo_return,
                'N/A',
                0 # placeholder for hqm scores
            ])
            
        except Exception as e:
            continue

    columns = [ #column names for the dataframe
        "Ticker",
        "Price",
        "Shares to Buy",
        "1y Price Return",
        "1y Return Percentile",
        "6mo Price Return",
        "6mo Return Percentile",
        "3mo Price Return",
        "3mo Return Percentile",
        "1mo Price Return",
        "1mo Return Percentile",
        "HQM Score"
        ]
    hqm_df = pd.DataFrame(momentum_stocks, columns=columns)

    time_periods = [ #time periods to loop through
        '1y',
        '6mo', 
        '3mo',
        '1mo' 
    ]

    # calculate percentiles and add to dataframe
    for time_period in time_periods:
        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        
        # remove NaN values from the series BEFORE calculating percentile
        clean_series = hqm_df[change_col].dropna()
        
        for row in hqm_df.index:
            single_value = hqm_df.loc[row, change_col]
            
            if pd.notna(single_value):  # only calculate if single value isn't NaN
                hqm_df.loc[row, percentile_col] = score(clean_series, single_value)
            else:
                hqm_df.loc[row, percentile_col] = 0  # placeholder

    for row in hqm_df.index:
        momentum_percentiles = []
        for time_period in time_periods:
            momentum_percentiles.append(hqm_df.loc[row, f'{time_period} Return Percentile'])
        hqm_df.loc[row, 'HQM Score'] = mean(momentum_percentiles) #takes the average of the percentiles for each stock, or "hqm" score

        # calculate volatility (risk) for each stock
    for row in hqm_df.index:
        returns = []
        for period in ['1y Price Return', '6mo Price Return', '3mo Price Return', '1mo Price Return']:
            if pd.notna(hqm_df.loc[row, period]):
                returns.append(hqm_df.loc[row, period])
        
        if len(returns) > 1:
            hqm_df.loc[row, 'Volatility'] = np.std(returns)
        else:
            hqm_df.loc[row, 'Volatility'] = 0.1  # default for limited data
    
    # create risk-adjusted score (higher = better risk/reward)
    for row in hqm_df.index:
        hqm_df.loc[row, 'Risk Adjusted Score'] = hqm_df.loc[row, 'HQM Score'] / (hqm_df.loc[row, 'Volatility'] + 0.01)

    
    # ACCURACY CHECK, manually checks to see if the scores were calculated correctly, notes: experienced timing error, hqm scores and percentiles were innacurate.
    print("\n" + "="*50)
    print("ACCURACY CHECK")
    print("="*50)
    
    # manually verify percentiles are properly calculated
    for time_period in time_periods:
        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        
        valid_data = hqm_df[hqm_df[change_col].notna()]
        if len(valid_data) > 0:
            highest_return_idx = valid_data[change_col].idxmax()
            lowest_return_idx = valid_data[change_col].idxmin()
            
            highest_percentile = hqm_df.loc[highest_return_idx, percentile_col]
            lowest_percentile = hqm_df.loc[lowest_return_idx, percentile_col]
            
            print(f"\n{time_period} Period Check:")
            print(f"  Highest return: {hqm_df.loc[highest_return_idx, change_col]:.4f} -> Percentile: {highest_percentile:.1f}")
            print(f"  Lowest return: {hqm_df.loc[lowest_return_idx, change_col]:.4f} -> Percentile: {lowest_percentile:.1f}")
            print(f"  ✓ Expected: Highest ~100, Lowest ~0")
    
    # verify HQM score calculation for a sample stock
    if len(hqm_df) > 0:
        sample_idx = hqm_df.index[0]  # First stock
        sample_ticker = hqm_df.loc[sample_idx, 'Ticker']
        
        print(f"\nHQM Score Verification for {sample_ticker}:")
        percentiles = []
        for time_period in time_periods:
            percentile_col = f'{time_period} Return Percentile'
            percentile_val = hqm_df.loc[sample_idx, percentile_col]
            percentiles.append(percentile_val)
            print(f"  {time_period} percentile: {percentile_val:.1f}")
        
        calculated_hqm = mean(percentiles)
        stored_hqm = hqm_df.loc[sample_idx, 'HQM Score']
        
        print(f"  Manual HQM calculation: {calculated_hqm:.2f}")
        print(f"  Stored HQM score: {stored_hqm:.2f}")
        print(f"  ✓ Match: {abs(calculated_hqm - stored_hqm) < 0.01}")
    
    print("="*50)
    # sorting the dataframe into something more tangible (top 50 hqm scores)
    hqm_df.sort_values("HQM Score", ascending = False, inplace = True) 
    # sort by risk-adjusted score (best risk/reward first)
    hqm_df.sort_values("Risk Adjusted Score", ascending=False, inplace=True)
    hqm_df = hqm_df[:20].reset_index(drop=True)  # top 20 stocks


def portfolio():
    global portfoliosz
    portfoliosz = input("Enter the size of your portfolio: ")
    try:
        float(portfoliosz)
    except ValueError:
        print('That is not a number!')
        print('Please try again:')
        portfoliosz = input("Enter the size of your portfolio: ")
    

    # calculate risk-adjusted position sizes
    total_risk_adj_score = hqm_df['Risk Adjusted Score'].sum()
    
    print(f"\nRisk-Adjusted Portfolio Allocation")
    print(f"Total Portfolio: ${float(portfoliosz):,.2f}")
    print(f"{'Rank':<4} {'Ticker':<7} {'HQM':<6} {'Risk':<6} {'Weight':<7} {'Shares':<7} {'Value'}")
    print("-" * 60)
    
    total_allocated = 0
    
    for i in range(0, len(hqm_df)):
        # weight based on risk-adjusted score (safer stocks get more money)
        risk_adj_score = hqm_df.loc[i, 'Risk Adjusted Score']
        weight = risk_adj_score / total_risk_adj_score
        position_value = float(portfoliosz) * weight
        shares_to_buy = math.floor(position_value / hqm_df.loc[i, "Price"])
        actual_value = shares_to_buy * hqm_df.loc[i, "Price"]
        
        # store calculated values
        hqm_df.loc[i, "Weight"] = weight
        hqm_df.loc[i, "Shares to Buy"] = shares_to_buy
        hqm_df.loc[i, "Position Value"] = actual_value
        
        total_allocated += actual_value
        
        # show allocation details
        ticker = hqm_df.loc[i, 'Ticker']
        hqm_score = hqm_df.loc[i, 'HQM Score']
        volatility = hqm_df.loc[i, 'Volatility']
        print(f"{i+1:<4} {ticker:<7} {hqm_score:<6.1f} {volatility:<6.3f} {weight:<7.1%} {shares_to_buy:<7} ${actual_value:>7,.0f}")
    
    # portfolio summary
    cash_remaining = float(portfoliosz) - total_allocated
    print("-" * 60)
    print(f"Total Invested: ${total_allocated:,.2f}")
    print(f"Cash Remaining: ${cash_remaining:,.2f} ({cash_remaining/float(portfoliosz):.1%})")
    
    # risk summary
    weighted_volatility = (hqm_df['Volatility'] * hqm_df['Weight']).sum()
    print(f"Portfolio Risk Score: {weighted_volatility:.3f} (lower is better)")


def xlsx_writer():
    writer = pd.ExcelWriter('momentum_strategy.xlsx', engine = 'xlsxwriter')
    hqm_df.to_excel(writer, sheet_name = 'Momentum Strategy', index = False)
    background_color = '#0a0a23'
    font_color = '#ffffff'
    # format for strings
    string_template = writer.book.add_format(
            {
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )
    # formats $ to the hundredths place
    dollar_template = writer.book.add_format(
            {
                'num_format':'$0.00',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )
    # only one decimal place to avoid partial stocks
    integer_template = writer.book.add_format(
            {
                'num_format':'0',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )
    # truncate to the hundredths place
    decimal_template = writer.book.add_format(
            {
                'num_format':'0.00',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )
    # adds % symbol, return is to the tenths place
    percent_template = writer.book.add_format(
            {
                'num_format':'0.0%',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )
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
        "N": ["Risk Adjusted Score", decimal_template],
        "O": ["Weight", percent_template],
        "P": ["Position Value", dollar_template]
    }
    for column in columns_format.keys():
        writer.sheets['Momentum Strategy'].set_column(f'{column}:{column}', 25, columns_format[column][1])
        writer.sheets['Momentum Strategy'].write(f'{column}1', columns_format[column][0], columns_format[column][1])
    writer.close()

def backtest():
    """
    Backtests the momentum strategy over 5 years, compares to S&P 500, 
    and projects returns through end of year
    """
    print("\n" + "="*60)
    print("5-YEAR MOMENTUM STRATEGY BACKTEST")
    print("="*60)
    
    # Get 5 years of S&P 500 data for benchmark
    spy = yf.Ticker("SPY")
    spy_hist = spy.history(period="5y")
    
    if len(spy_hist) < 1000:  # Need substantial data
        print("Insufficient data for 5-year backtesting")
        return
    
    # Calculate S&P 500 5-year performance
    spy_start = spy_hist['Close'].iloc[0]
    spy_end = spy_hist['Close'].iloc[-1]
    spy_total_return = (spy_end - spy_start) / spy_start
    spy_annual_return = (1 + spy_total_return) ** (1/5) - 1
    
    print(f"Backtesting period: {spy_hist.index[0].date()} to {spy_hist.index[-1].date()}")
    print(f"S&P 500 Benchmark:")
    print(f"  5-Year Total Return: {spy_total_return:.1%}")
    print(f"  Annualized Return: {spy_annual_return:.1%}")
    
    # Strategy backtesting - simulate rebalancing annually
    strategy_returns = []
    years_data = []
    
    for year in range(5):
        year_start_idx = int(len(spy_hist) * year / 5)
        year_end_idx = int(len(spy_hist) * (year + 1) / 5) - 1
        
        if year_end_idx >= len(spy_hist):
            year_end_idx = len(spy_hist) - 1
        
        # For each year, test our top momentum stocks
        year_return = 0
        valid_stocks = 0
        
        for i in range(min(20, len(hqm_df))):
            ticker = hqm_df.iloc[i]['Ticker']
            
            try:
                stock_data = yf.Ticker(ticker)
                stock_hist = stock_data.history(period="5y")
                
                if len(stock_hist) < year_end_idx + 10:
                    continue
                
                # Calculate this stock's return for the year
                start_price = stock_hist['Close'].iloc[year_start_idx]
                end_price = stock_hist['Close'].iloc[year_end_idx]
                stock_year_return = (end_price - start_price) / start_price
                
                year_return += stock_year_return
                valid_stocks += 1
                
            except:
                continue
        
        if valid_stocks > 0:
            avg_year_return = year_return / valid_stocks
            strategy_returns.append(avg_year_return)
            years_data.append({
                'year': year + 1,
                'return': avg_year_return,
                'valid_stocks': valid_stocks
            })
    
    if len(strategy_returns) == 0:
        print("Unable to calculate strategy returns")
        return
    
    # Calculate strategy performance metrics
    strategy_total_return = 1
    for annual_return in strategy_returns:
        strategy_total_return *= (1 + annual_return)
    strategy_total_return -= 1
    
    strategy_annual_return = (1 + strategy_total_return) ** (1/5) - 1
    strategy_volatility = np.std(strategy_returns)
    
    # Performance comparison
    print(f"\nMOMENTUM STRATEGY RESULTS:")
    print(f"  5-Year Total Return: {strategy_total_return:.1%}")
    print(f"  Annualized Return: {strategy_annual_return:.1%}")
    print(f"  Annual Volatility: {strategy_volatility:.1%}")
    
    excess_return = strategy_annual_return - spy_annual_return
    print(f"\nPERFORMANCE vs S&P 500:")
    print(f"  Excess Annual Return: {excess_return:+.1%}")
    print(f"  Outperformed S&P 500: {'YES' if excess_return > 0 else 'NO'}")
    
    # Year-by-year breakdown
    print(f"\nYEAR-BY-YEAR PERFORMANCE:")
    print(f"{'Year':<6} {'Strategy':<12} {'Stocks Used'}")
    print("-" * 30)
    for year_data in years_data:
        print(f"{year_data['year']:<6} {year_data['return']:<11.1%} {year_data['valid_stocks']}")
    
    # project returns from oct 1 2025 to oct 1 2026
    from datetime import datetime
    
    oct_1_2025 = datetime(2025, 10, 1)
    oct_1_2026 = datetime(2026, 10, 1)
    current_date = datetime.now()
    
    # calculate how much of the projection period has passed
    if current_date > oct_1_2025:
        days_passed = (current_date - oct_1_2025).days
        days_total = (oct_1_2026 - oct_1_2025).days
        fraction_remaining = (days_total - days_passed) / days_total
    else:
        fraction_remaining = 1.0  # full year projection
    
    projected_return = strategy_annual_return * fraction_remaining
    
    print(f"Projected return (Oct 1 2025 - Oct 1 2026): {projected_return:.1%}")
    
    # final summary statement
    excess_return = strategy_annual_return - spy_annual_return
    print(f"\nStrategy outperformed S&P 500 by {excess_return:+.1%} over the 5-year period (2020-2025), expected to return {projected_return:.1%} from Oct 1 2025 to Oct 1 2026.")
    
    return {
        'strategy_annual_return': strategy_annual_return,
        'strategy_total_return': strategy_total_return,
        'spy_annual_return': spy_annual_return,
        'excess_return': excess_return,
        'outperformed': excess_return > 0
    }



def main():
    datafetcher()
    portfolio()
    xlsx_writer()
    backtest()
main()