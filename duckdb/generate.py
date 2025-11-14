import duckdb
import time

# Configuration
NUM_STOCKS = 500
TRADING_DAYS = 2500  # roughly 10 years
TOTAL_ROWS = NUM_STOCKS * TRADING_DAYS


print(f"Generating {TOTAL_ROWS:,} synthetic stock price records...")
print(f"Stocks: {NUM_STOCKS}, Trading days: {TRADING_DAYS}")
print(f"Expected dataset size: ~{TOTAL_ROWS / 1_000_000:.1f}M rows\n")

# Connect to DuckDB
conn = duckdb.connect("stocks.duckdb")

# Generate data directly in DuckDB
print("Creating synthetic data in DuckDB...")
conn.execute(
    """
    CREATE TABLE stock_data AS
    SELECT
        (r % ?) as stock_id,
        'STOCK' || LPAD(CAST((r % ?) as VARCHAR), 3, '0') as symbol,
        DATE '2014-01-02' + CAST((r / ?) as INT) as date,
        100 + (r * 73 % 100) as close,
        1000000 + (r * 83 % 9000000) as volume
    FROM (SELECT row_number() OVER () - 1 as r FROM range(?))
""",
    [NUM_STOCKS, NUM_STOCKS, NUM_STOCKS, TOTAL_ROWS],
)

print(f"Data created. Verifying dataset...\n")

# Verify dataset
verify = conn.execute(
    """
    SELECT
        COUNT(*) as total_records,
        COUNT(DISTINCT symbol) as unique_stocks,
        COUNT(DISTINCT date) as unique_dates,
        MIN(date) as earliest_date,
        MAX(date) as latest_date
    FROM stock_data
"""
).fetchall()

print(f"Total records: {verify[0][0]:,}")
print(f"Unique stocks: {verify[0][1]}")
print(f"Unique dates: {verify[0][2]}")
print(f"Date range: {verify[0][3]} to {verify[0][4]}\n")

conn.close()
