# DuckDB Excel Add-in Demo

A simple demo of using [xllify](https://xllify.com) to bring DuckDB's analytical capabilities into Excel. This add-in generates 1.25M rows of synthetic stock market data in-memory and provides pre-built SQL queries as Excel functions.

## Overview

This example shows:

- **In-memory analytics**: 500 stocks with 2,500 trading days each (10 years of synthetic data)
- **DuckDB integration**: Leverage DuckDB's fast analytical query engine from Excel
- **Matrix results**: Query results are returned as dynamic arrays that spill into Excel cells
- **Window functions**: Advanced SQL analytics including moving averages, rankings, and percentiles

## Quick Start

### Installation

```bash
# Create and activate a virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Build the Excel add-in
xllify duck.xll duckdb_stock_demo.py --py-requirements requirements.txt
```

### Usage

1. Open `duck.xll` in Excel
2. Create a new worksheet
3. Start typing `=duckdb.` to see all available functions
4. Try an example:
   ```excel
   =duckdb.ComplexAggregation()
   ```

The function will return a dynamic array with stock statistics including averages, volatility, and trading volumes.

![demo](./demo.png)

## Available Functions

All functions return results as Excel dynamic arrays (spilling into multiple cells):

### `duckdb.MovingAvg()`

Calculate 20-day and 50-day moving averages for stock prices. Includes day numbering to track position in the time series.

### `duckdb.QuarterlyPerformance()`

Rank stocks by quarterly performance, showing the top 3 performers for each quarter. Demonstrates RANK window function with partitioning.

### `duckdb.PriceChanges()`

Analyze daily price changes using LAG/LEAD window functions. Shows largest price movements with cumulative returns from the first trading day.

### `duckdb.ComplexAggregation()`

Aggregate statistics per stock including 30-day moving averages, ordered by volatility (standard deviation).

### `duckdb.NTILEDistribution()`

Distribute monthly returns into quartiles using NTILE. Shows how returns are distributed across different performance buckets.

### `duckdb.MonthlyVolatility()`

Calculate monthly price volatility and track how it changes over time. Useful for identifying periods of market turbulence.

### `duckdb.PercentileRanking()`

Rank stocks using multiple ranking functions (PERCENT_RANK, DENSE_RANK, ROW_NUMBER) to identify top performers across different time windows.

## Data schema

The in-memory database contains a single table `stock_data` with the following structure:

| Column   | Type    | Description                        |
| -------- | ------- | ---------------------------------- |
| stock_id | INTEGER | Numeric stock identifier (0-499)   |
| symbol   | VARCHAR | Stock symbol (STOCK000-STOCK499)   |
| date     | DATE    | Trading date (2014-01-02 to ~2023) |
| close    | DECIMAL | Closing price                      |
| volume   | INTEGER | Trading volume                     |

Total records: **1,250,000** (500 stocks × 2,500 trading days)

## How It Works

1. **Startup**: When the add-in loads, it creates an in-memory DuckDB database and generates synthetic stock data
2. **Query Execution**: Each Excel function executes a pre-defined SQL query against the in-memory database
3. **Result Formatting**: Query results are converted to a 2D matrix with headers and returned to Excel as dynamic arrays
4. **Performance**: DuckDB's columnar storage and vectorized execution enable fast queries even on 1M+ rows

## Approach

- **[DuckDB](https://duckdb.org/)**: Analytical database with excellent SQL support
- **[xllify](https://xllify.com)**: Converts Python functions into Excel add-ins
- **Window Functions**: SQL features for analytics (moving averages, rankings, percentiles)
- **Dynamic Arrays**: Excel feature that allows functions to return multiple values

## Customization

To add your own queries:

1. Define a new function in `duckdb_stock_demo.py`:

   ```python
   @xllify.fn("duckdb.MyQuery")
   def duckdb_MyQuery():
       """Your query description."""
       query = """
       SELECT * FROM stock_data
       WHERE symbol = 'STOCK001'
       LIMIT 10
       """
       return get_results_matrix(query, conn)
   ```

2. Rebuild the add-in:

   ```bash
   xllify duck.xll duckdb_stock_demo.py --py-requirements requirements.txt
   ```

3. Reload the add-in in Excel

## Performance

- First load generates 1.25M rows which takes 1-2 seconds
- Queries typically execute in under 100ms
- Resultset size is limited to prevent overwhelming Excel (configurable in code)
- All data is in-memory and lost when Excel closes (that's fine, it's only a demo)
- Data is naively cached after first query run, so use `xllify-clear-cache` to re-run

## Links

- [xllify Documentation](https://xllify.com)
- [DuckDB Documentation](https://duckdb.org/docs/)
- [Excel Dynamic Arrays](https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
