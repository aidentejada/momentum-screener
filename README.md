# momentum-screener
# Quantitative Equity Strategy

A momentum-based stock selection system that screens S&P 500 stocks using multi-timeframe analysis and constructs portfolios using CAPM expected returns and volatility-adjusted position sizing.

## Overview

This strategy selects the top 20 stocks from the S&P 500 based on momentum performance across multiple time periods, then allocates portfolio weights using expected returns and risk metrics.

## Strategy Components

### Momentum Analysis
The system calculates returns across four time periods:
- 1-year return (25% weight)
- 6-month return (25% weight) 
- 3-month return (30% weight)
- 1-month return (20% weight)

Each stock receives percentile rankings for each period, which are combined into a High Quality Momentum (HQM) score.

### Expected Return Calculation
Uses CAPM model to estimate expected returns:
```
Expected Return = Risk-Free Rate + Beta × (Market Return - Risk-Free Rate)
```
- Risk-free rate: 4.1%
- Market return assumption: 9.5%
- Beta sourced from Yahoo Finance

### Position Sizing
Portfolio weights calculated using:
```
Weight = (HQM Score × Expected Return) ÷ (Volatility + 0.05)
```

Risk management rules:
- Maximum 6.5% position size
- Volatility capped at 13% for calculations
- Mega-cap stocks (AAPL, MSFT, GOOG, GOOGL, NVDA, META) receive 15% liquidity premium

### Portfolio Construction
- Selects top 20 stocks by HQM score
- Calculates share quantities (rounded down to whole shares)
- Normalizes weights to sum to 100%
- Reports cash remaining from rounding

## Code Structure

**datafetcher()**: Retrieves stock data, calculates momentum metrics, percentile rankings, and position weights

**portfolio()**: Takes portfolio size input, calculates share quantities, displays allocation table with risk metrics

**xlsx_writer()**: Exports results to formatted Excel spreadsheet with professional styling

## Requirements

- Python packages: pandas, numpy, yfinance, scipy, xlsxwriter
- S&P 500 stock list (sp500.csv file)
- Internet connection for real-time price data

## Output

The system generates:
- Console display of top 20 stock allocations
- Portfolio expected return and risk score
- Excel export with all calculated metrics
- Cash allocation and remaining balance

## Usage

Run the script and enter your portfolio size when prompted. The system will fetch current data, perform calculations, and display/export results.
