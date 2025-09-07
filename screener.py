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
    hqm_df = hqm_df[:50].reset_index(drop = True)


def portfolio():
    global portfoliosz
    portfoliosz = input("Enter the size of your portfolio: ")

    try:
        float(portfoliosz)
    except ValueError:
        print('That is not a number!')
        print('Please try again:')
        portfoliosz = input("Enter the size of your portfolio: ")
    pos_sz = float(portfoliosz)/len(hqm_df.index) # calculate position size
    for i in range(0, len(hqm_df)):
        hqm_df.loc[i, "Shares to Buy"] = math.floor(pos_sz/hqm_df.loc[i, "Price"]) # calculates shares to buy



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
        "L": ["HQM Score", decimal_template]
    }
    for column in columns_format.keys():
        writer.sheets['Momentum Strategy'].set_column(f'{column}:{column}', 25, columns_format[column][1])
        writer.sheets['Momentum Strategy'].write(f'{column}1', columns_format[column][0], columns_format[column][1])
    writer.close()
datafetcher()
portfolio()
xlsx_writer()