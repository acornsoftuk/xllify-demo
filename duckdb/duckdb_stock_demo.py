import duckdb
import xllify

# Connect to DuckDB
conn = duckdb.connect(":memory:")

# Configuration
NUM_STOCKS = 500
TRADING_DAYS = 2500  # roughly 10 years
TOTAL_ROWS = NUM_STOCKS * TRADING_DAYS

print(f"Generating {TOTAL_ROWS:,} synthetic stock price records...")
print(f"Stocks: {NUM_STOCKS}, Trading days: {TRADING_DAYS}")
print(f"Expected dataset size: ~{TOTAL_ROWS / 1_000_000:.1f}M rows\n")

# Generate fake data directly in DuckDB in-memory
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

def get_results_matrix(query, conn, limit_rows=None):
    """Helper to execute query and return results as 2D matrix with headers"""
    result = conn.execute(query).fetchall()

    # Build the 2D matrix: [header_row, data_rows...]
    matrix = [[desc[0] for desc in conn.execute(query).description]]

    # Add data rows (optionally limited)
    if limit_rows is not None:
        matrix.extend(list(row) for row in result[:limit_rows])
    else:
        matrix.extend(list(row) for row in result)

    return matrix


@xllify.fn("duckdb.MovingAvg")
def duckdb_MovingAvg():
    """Calculate 20-day and 50-day moving averages for stock prices with day numbering."""
    query = """
    SELECT
        symbol,
        date,
        close,
        ROUND(AVG(close) OVER (
            PARTITION BY symbol 
            ORDER BY date 
            ROWS BETWEEN 19 PRECEDING AND CURRENT ROW
        ), 2) as ma_20,
        ROUND(AVG(close) OVER (
            PARTITION BY symbol 
            ORDER BY date 
            ROWS BETWEEN 49 PRECEDING AND CURRENT ROW
        ), 2) as ma_50,
        ROW_NUMBER() OVER (PARTITION BY symbol ORDER BY date) as day_num
    FROM stock_data
    WHERE symbol IN (SELECT DISTINCT symbol FROM stock_data LIMIT 5)
    LIMIT 30
    """
    return get_results_matrix(query, conn, limit_rows=100)


@xllify.fn("duckdb.QuarterlyPerformance")
def duckdb_QuarterlyPerformance():
    """Rank stocks by quarterly performance, showing top 3 performers per quarter."""
    query = """
    WITH quarterly AS (
        SELECT
            DATE_TRUNC('quarter', date)::DATE as quarter,
            symbol,
            COUNT(*) as trading_days,
            ROUND(100.0 * (MAX(close) - MIN(close)) / NULLIF(MIN(close), 0), 2) as quarterly_return,
            RANK() OVER (
                PARTITION BY DATE_TRUNC('quarter', date)
                ORDER BY MAX(close) - MIN(close) DESC
            ) as return_rank
        FROM stock_data
        GROUP BY DATE_TRUNC('quarter', date), symbol
    )
    SELECT
        quarter,
        symbol,
        trading_days,
        quarterly_return,
        return_rank
    FROM quarterly
    WHERE return_rank <= 3
    ORDER BY quarter DESC, return_rank
    LIMIT 30
    """
    return get_results_matrix(query, conn, limit_rows=100)


@xllify.fn("duckdb.PriceChanges")
def duckdb_PriceChanges():
    """Analyze daily price changes using LAG/LEAD window functions, showing largest movements."""
    query = """
    WITH price_analysis AS (
        SELECT
            symbol,
            date,
            close,
            LAG(close) OVER (PARTITION BY symbol ORDER BY date) as prev_close,
            LEAD(close) OVER (PARTITION BY symbol ORDER BY date) as next_close,
            FIRST_VALUE(close) OVER (PARTITION BY symbol ORDER BY date) as first_price,
            LAST_VALUE(close) OVER (PARTITION BY symbol ORDER BY date
                ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING) as last_price
        FROM stock_data
    )
    SELECT
        symbol,
        date,
        close,
        ROUND(COALESCE(close - prev_close, 0), 2) as daily_change,
        ROUND(100.0 * (close - first_price) / NULLIF(first_price, 0), 2) as cumulative_return
    FROM price_analysis
    WHERE daily_change != 0
    ORDER BY ABS(daily_change) DESC
    LIMIT 20
    """
    return get_results_matrix(query, conn, limit_rows=100)


@xllify.fn("duckdb.ComplexAggregation")
def duckdb_ComplexAggregation():
    """Aggregate stock statistics with 30-day moving averages, ordered by volatility."""
    query = """
    WITH daily_stats AS (
        SELECT
            symbol,
            date,
            close,
            volume,
            AVG(close) OVER (PARTITION BY symbol ORDER BY date ROWS BETWEEN 29 PRECEDING AND CURRENT ROW) as ma_30,
            ROW_NUMBER() OVER (PARTITION BY symbol ORDER BY date) as day_num
        FROM stock_data
    )
    SELECT
        symbol,
        COUNT(*) as total_days,
        ROUND(AVG(close), 2) as avg_close,
        ROUND(MAX(close), 2) as max_close,
        ROUND(MIN(close), 2) as min_close,
        ROUND(STDDEV_POP(close), 2) as std_dev,
        ROUND(AVG(volume), 0) as avg_volume
    FROM daily_stats
    GROUP BY symbol
    ORDER BY std_dev DESC
    LIMIT 20
    """
    return get_results_matrix(query, conn, limit_rows=100)


@xllify.fn("duckdb.NTILEDistribution")
def duckdb_NTILEDistribution():
    """Distribute monthly returns into quartiles using NTILE, showing distribution statistics."""
    query = """
    WITH monthly_returns AS (
        SELECT
            symbol,
            DATE_TRUNC('month', date)::DATE as month,
            ROUND(100.0 * (MAX(close) - MIN(close)) / NULLIF(MIN(close), 0), 2) as monthly_return
        FROM stock_data
        GROUP BY symbol, DATE_TRUNC('month', date)
    ),
    returns_with_ntile AS (
        SELECT
            month,
            monthly_return,
            NTILE(4) OVER (PARTITION BY month ORDER BY monthly_return) as return_quartile
        FROM monthly_returns
    )
    SELECT
        month,
        return_quartile,
        COUNT(*) as stock_count,
        ROUND(MIN(monthly_return), 2) as min_return,
        ROUND(AVG(monthly_return), 2) as avg_return,
        ROUND(MAX(monthly_return), 2) as max_return
    FROM returns_with_ntile
    GROUP BY month, return_quartile
    ORDER BY month DESC, return_quartile
    LIMIT 20
    """
    return get_results_matrix(query, conn, limit_rows=100)


@xllify.fn("duckdb.MonthlyVolatility")
def duckdb_MonthlyVolatility():
    """Calculate monthly price volatility and track changes over time."""
    query = """
    WITH daily_volatility AS (
        SELECT
            symbol,
            DATE_TRUNC('month', date)::DATE as month,
            close,
            AVG(close) OVER (PARTITION BY symbol, DATE_TRUNC('month', date)) as monthly_avg
        FROM stock_data
    ),
    monthly_stats AS (
        SELECT
            symbol,
            month,
            ROUND(SQRT(AVG((close - monthly_avg) * (close - monthly_avg))), 2) as monthly_volatility,
            ROUND(AVG(close), 2) as avg_price,
            ROW_NUMBER() OVER (PARTITION BY symbol ORDER BY month) as month_num
        FROM daily_volatility
        GROUP BY symbol, month, monthly_avg
    )
    SELECT
        symbol,
        month,
        monthly_volatility,
        avg_price,
        LAG(monthly_volatility) OVER (PARTITION BY symbol ORDER BY month) as prev_volatility,
        ROUND(monthly_volatility - LAG(monthly_volatility) OVER (PARTITION BY symbol ORDER BY month), 2) as vol_change
    FROM monthly_stats
    ORDER BY month DESC, monthly_volatility DESC
    LIMIT 25
    """
    return get_results_matrix(query, conn, limit_rows=100)


@xllify.fn("duckdb.PercentileRanking")
def duckdb_PercentileRanking():
    """Rank stocks using percentile, dense rank, and row number functions to identify top performers."""
    query = """
    WITH daily_data AS (
        SELECT
            symbol,
            date,
            close,
            PERCENT_RANK() OVER (PARTITION BY symbol ORDER BY close) as price_percentile,
            DENSE_RANK() OVER (ORDER BY close DESC) as global_price_rank,
            ROW_NUMBER() OVER (PARTITION BY DATE_TRUNC('month', date) ORDER BY close DESC) as monthly_price_rank
        FROM stock_data
    )
    SELECT
        symbol,
        date,
        ROUND(close, 2) as close,
        ROUND(price_percentile * 100, 2) as price_percentile,
        global_price_rank,
        monthly_price_rank
    FROM daily_data
    WHERE global_price_rank <= 10
    ORDER BY global_price_rank
    LIMIT 15
    """
    return get_results_matrix(query, conn, limit_rows=100)
