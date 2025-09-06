###########################
import numpy as np        #
import pandas as pd       #
import math               #
from scipy import stats   #
import xlsxwriter         #
import yfinance as yf     #
###########################

def datafetcher():
    global final_df
    # import list of stocks
    stocks = pd.read_csv("sp500.csv")
    momentum_stocks = []


    for stock in stocks['Symbol']:
        try:
            ticker = yf.Ticker(stock)
            hist = ticker.history(period='1y')
            
            if len(hist) == 0:
                continue
                
            current_price = hist['Close'].iloc[-1]
            price_1yr_ago = hist['Close'].iloc[0] 
            one_year_return = (current_price - price_1yr_ago) / price_1yr_ago
            
            momentum_stocks.append([stock, current_price, one_year_return, 0])
            
        except Exception as e:
            continue

    columns = ["Ticker", "Price", "1y Price Return", "Shares to buy"]
    final_df = pd.DataFrame(momentum_stocks, columns=columns)
    
    final_df.sort_values('1y Price Return', ascending = False, inplace = True)
    final_df = final_df[:50].reset_index()

def portfolio():
    global portfoliosz
    portfoliosz = input("Enter the size of your portfolio: ")

    try:
        float(portfoliosz)
    except ValueError:
        print('That is not a number!')
        print('Please try again:')
        portfoliosz = input("Enter the size of your portfolio: ")
datafetcher()
portfolio()
print(portfoliosz)

pos_sz = float(portfoliosz)/len(final_df.index)
for i in range(0, len(final_df)):
    final_df.loc[i, "Shares to buy"] = math.floor(pos_sz/final_df.loc[i, "Price"])
print(final_df)
