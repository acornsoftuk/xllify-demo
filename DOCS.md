# xllify-demo Function Documentation

This document provides comprehensive documentation for all custom Excel functions implemented in this repository using [xllify.com](https://xllify.com).

## Table of Contents

- [Simple Demo Functions](#simple-demo-functions)
- [Black-Scholes Option Pricing](#black-scholes-option-pricing)
- [Portfolio Analytics](#portfolio-analytics)
- [Market Simulation](#market-simulation)
- [Trading Analytics](#trading-analytics)
- [Technical Analysis](#technical-analysis)
- [Fixed Income](#fixed-income)
- [Risk Management](#risk-management)
- [Option Greeks](#option-greeks)
- [Demo Data Generators](#demo-data-generators)

---

## Simple Demo Functions

### xllify.Demo.Hello

**File:** `simple.luau:1`

**Description:** Says hello!

**Category:** xllify Demos

**Parameters:**

- `name` (string): Name to greet

**Examples:**

```excel
=xllify.Demo.Hello("Biddy")
// Returns: "Hello, Biddy!"

=xllify.Demo.Hello()
// Returns: "Hello, World!"
```

---

### xllify.Demo.AgeCategory

**File:** `simple.luau:13`

**Description:** Determines your life stage based on age

**Category:** xllify Demos

**Parameters:**

- `age` (number): Age in years

**Examples:**

```excel
=xllify.Demo.AgeCategory(5)
// Returns: "Kid"

=xllify.Demo.AgeCategory(17)
// Returns: "Annoying teenager"

=xllify.Demo.AgeCategory(42)
// Returns: "Disillusioned adult"

=xllify.Demo.AgeCategory(70)
// Returns: "OK boomer"
```

**Age Categories:**

- Less than 0: "Baby"
- 0-12: "Kid"
- 13-19: "Annoying teenager"
- 20-64: "Disillusioned adult"
- 65+: "OK boomer"

---

### xllify.Demo.Portfolio

**File:** `simple.luau:34`

**Description:** Generate sample portfolio holdings

**Category:** xllify Demos

**Parameters:**

- `num_positions` (number): Number of positions to generate in the portfolio (default: 10)
- `total_value` (number): Total portfolio value to distribute across positions (default: 1,000,000)
- `seed` (number): Random seed for reproducible portfolio generation (default: 321)

**Returns:** A 2D array with columns: Ticker, Quantity, Price, Total Value

**Examples:**

```excel
// Generate a 10-position portfolio worth $1M with default seed
=xllify.Demo.Portfolio(10, 1000000, 321)

// Generate a 5-position portfolio worth $500K
=xllify.Demo.Portfolio(5, 500000)

// Use default parameters
=xllify.Demo.Portfolio()

// Generate portfolio with custom seed for different data
=xllify.Demo.Portfolio(8, 750000, 999)
```

**Output Format:**

The function returns a dynamic array with 4 columns:

| Ticker | Quantity | Price  | Total Value |
| ------ | -------- | ------ | ----------- |
| AAPL   | 2456     | 178.45 | 438,233.20  |
| MSFT   | 1823     | 215.67 | 393,166.41  |
| ...    | ...      | ...    | ...         |

**Notes:**

- Uses tickers from major US companies (AAPL, MSFT, GOOGL, AMZN, META, etc.)
- Prices are randomly generated between $50-$250
- The same seed value will always produce the same portfolio
- Quantities are calculated to distribute the total portfolio value across positions

---

## Black-Scholes Option Pricing

### bs.Call

**File:** `black_scholes.luau:21`

**Description:** Black-Scholes call option price

**Parameters:**

- `S` (number): Current stock price
- `K` (number): Strike price
- `T` (number): Time to expiration (years)
- `r` (number): Risk-free interest rate
- `sigma` (number): Volatility (annualized)

**Returns:** Call option price

**Examples:**

```excel
// Price a call option with spot=100, strike=100, 1 year, 5% rate, 20% vol
=bs.Call(100, 100, 1, 0.05, 0.20)
```

---

### bs.Put

**File:** `black_scholes.luau:37`

**Description:** Black-Scholes put option price

**Parameters:**

- `S` (number): Current stock price
- `K` (number): Strike price
- `T` (number): Time to expiration (years)
- `r` (number): Risk-free interest rate
- `sigma` (number): Volatility (annualized)

**Returns:** Put option price

**Examples:**

```excel
// Price a put option with spot=100, strike=100, 1 year, 5% rate, 20% vol
=bs.Put(100, 100, 1, 0.05, 0.20)
```

**Notes:**

- Both functions use the standard Black-Scholes formula
- When T <= 0, returns intrinsic value only
- Uses a normal CDF approximation for pricing

---

## Portfolio Analytics

### Portfolio.Returns

**File:** `trading.luau:10`

**Description:** Calculate returns from a price series

**Category:** Portfolio Analytics

**Parameters:**

- `prices` (array): Array of prices

**Returns:** Array of returns (length = prices length - 1)

**Examples:**

```excel
// Calculate returns from price series in A1:A100
=Portfolio.Returns(A1:A100)
```

**Notes:**

- Handles both 1D and 2D arrays
- Returns are calculated as: (price[i] - price[i-1]) / price[i-1]
- Returns 0 for invalid or non-numeric values

---

### Portfolio.Volatility

**File:** `trading.luau:35`

**Description:** Calculate annualized volatility from returns

**Category:** Portfolio Analytics

**Parameters:**

- `returns` (array): Array of returns
- `periods_per_year` (number): Number of periods per year (default: 252 for daily data)

**Returns:** Annualized volatility as a decimal

**Examples:**

```excel
// Calculate annualized volatility from daily returns
=Portfolio.Volatility(A1:A100, 252)

// Calculate from monthly returns
=Portfolio.Volatility(A1:A12, 12)
```

**Notes:**

- Uses sample standard deviation (n-1 denominator)
- Automatically annualizes based on periods_per_year parameter

---

### Portfolio.Sharpe

**File:** `trading.luau:78`

**Description:** Calculate Sharpe ratio (excess return / volatility)

**Category:** Portfolio Analytics

**Parameters:**

- `returns` (array): Array of returns
- `risk_free_rate` (number): Annual risk-free rate (default: 0.02 = 2%)
- `periods_per_year` (number): Number of periods per year (default: 252)

**Returns:** Sharpe ratio

**Examples:**

```excel
// Calculate Sharpe ratio with 2% risk-free rate
=Portfolio.Sharpe(A1:A100, 0.02, 252)

// Use default risk-free rate
=Portfolio.Sharpe(A1:A100)
```

---

## Market Simulation

### Simulation.PricePath

**File:** `trading.luau:140`

**Description:** Simulate geometric Brownian motion price path

**Category:** Market Simulation

**Parameters:**

- `S0` (number): Initial stock price (default: 100)
- `mu` (number): Annual drift (default: 0.10 = 10%)
- `sigma` (number): Annual volatility (default: 0.20 = 20%)
- `days` (number): Number of days to simulate (default: 252)
- `seed` (number): Random seed for reproducibility (default: 42)

**Returns:** Array of simulated prices

**Examples:**

```excel
// Simulate 1 year of daily prices starting at 100
=Simulation.PricePath(100, 0.10, 0.20, 252, 42)

// Use default parameters
=Simulation.PricePath()
```

**Notes:**

- Uses geometric Brownian motion: dS = μS dt + σS dW
- Same seed produces identical results (deterministic)
- Daily timestep (dt = 1/252)

---

### Simulation.MonteCarlo

**File:** `trading.luau:166`

**Description:** Run Monte Carlo option pricing simulation

**Category:** Market Simulation

**Parameters:**

- `S` (number): Current stock price (default: 100)
- `K` (number): Strike price (default: 100)
- `T` (number): Time to expiration in years (default: 1)
- `r` (number): Risk-free rate (default: 0.05)
- `sigma` (number): Volatility (default: 0.20)
- `num_sims` (number): Number of simulations (default: 10000)
- `option_type` (string): "call" or "put" (default: "call")

**Returns:** Estimated option price

**Examples:**

```excel
// Price a call option with 10000 simulations
=Simulation.MonteCarlo(100, 100, 1, 0.05, 0.20, 10000, "call")

// Price a put option
=Simulation.MonteCarlo(100, 100, 1, 0.05, 0.20, 10000, "put")
```

---

## Trading Analytics

### Trading.VWAP

**File:** `trading.luau:205`

**Description:** Calculate volume-weighted average price

**Category:** Trading Analytics

**Parameters:**

- `prices` (array): Array of prices
- `volumes` (array): Array of volumes

**Returns:** VWAP

**Examples:**

```excel
// Calculate VWAP from prices and volumes
=Trading.VWAP(A1:A100, B1:B100)
```

**Formula:**

```
VWAP = Σ(price[i] × volume[i]) / Σ(volume[i])
```

---

### Trading.Slippage

**File:** `trading.luau:236`

**Description:** Calculate execution slippage vs benchmark

**Category:** Trading Analytics

**Parameters:**

- `benchmark_price` (number): Benchmark price (e.g., arrival price)
- `execution_prices` (array): Array of execution prices
- `quantities` (array): Array of quantities

**Returns:** Slippage as a percentage (positive = paid more, negative = saved)

**Examples:**

```excel
// Calculate slippage vs benchmark of 100
=Trading.Slippage(100, A1:A10, B1:B10)
```

**Notes:**

- Returns slippage as decimal (e.g., 0.001 = 0.1% slippage)
- Useful for analyzing execution quality

---

## Technical Analysis

### Technical.SMA

**File:** `trading.luau:271`

**Description:** Simple moving average

**Category:** Technical Analysis

**Parameters:**

- `prices` (array): Array of prices
- `window` (number): Window size (default: 20)

**Returns:** Array of SMA values

**Examples:**

```excel
// Calculate 20-period SMA
=Technical.SMA(A1:A100, 20)

// Calculate 50-period SMA
=Technical.SMA(A1:A100, 50)
```

**Notes:**

- Returns array of length: (prices length - window + 1)
- First value is the SMA of the first 'window' prices

---

### Technical.EMA

**File:** `trading.luau:307`

**Description:** Exponential moving average

**Category:** Technical Analysis

**Parameters:**

- `prices` (array): Array of prices
- `window` (number): Window size (default: 20)

**Returns:** Array of EMA values

**Examples:**

```excel
// Calculate 20-period EMA
=Technical.EMA(A1:A100, 20)
```

**Notes:**

- Uses smoothing factor: α = 2 / (window + 1)
- First EMA value is initialized with SMA

---

### Technical.RSI

**File:** `trading.luau:350`

**Description:** Relative Strength Index

**Category:** Technical Analysis

**Parameters:**

- `prices` (array): Array of prices
- `period` (number): RSI period (default: 14)

**Returns:** Array of RSI values (0-100)

**Examples:**

```excel
// Calculate 14-period RSI
=Technical.RSI(A1:A100, 14)
```

**Notes:**

- RSI values range from 0 to 100
- Typically: RSI > 70 = overbought, RSI < 30 = oversold
- Uses Wilder's smoothing method

---

## Fixed Income

### Bond.Price

**File:** `trading.luau:430`

**Description:** Calculate bond price from yield

**Category:** Fixed Income

**Parameters:**

- `face_value` (number): Face value of bond (default: 1000)
- `coupon_rate` (number): Annual coupon rate (default: 0.05 = 5%)
- `ytm` (number): Yield to maturity (default: 0.06 = 6%)
- `years` (number): Years to maturity (default: 10)
- `frequency` (number): Coupon frequency per year (default: 2 = semi-annual)

**Returns:** Bond price

**Examples:**

```excel
// Price a 10-year bond with 5% coupon, 6% YTM
=Bond.Price(1000, 0.05, 0.06, 10, 2)
```

**Notes:**

- When YTM > coupon rate, price < face value (discount bond)
- When YTM < coupon rate, price > face value (premium bond)

---

### Bond.Duration

**File:** `trading.luau:458`

**Description:** Calculate Macaulay duration

**Category:** Fixed Income

**Parameters:**

- `face_value` (number): Face value of bond (default: 1000)
- `coupon_rate` (number): Annual coupon rate (default: 0.05)
- `ytm` (number): Yield to maturity (default: 0.06)
- `years` (number): Years to maturity (default: 10)
- `frequency` (number): Coupon frequency per year (default: 2)

**Returns:** Macaulay duration in years

**Examples:**

```excel
// Calculate duration of a 10-year bond
=Bond.Duration(1000, 0.05, 0.06, 10, 2)
```

**Notes:**

- Duration measures bond price sensitivity to yield changes
- Higher duration = higher interest rate risk

---

## Risk Management

### Risk.VaR

**File:** `trading.luau:495`

**Description:** Calculate Value at Risk (historical method)

**Category:** Risk Management

**Parameters:**

- `returns` (array): Array of returns
- `confidence_level` (number): Confidence level (default: 0.95 = 95%)

**Returns:** VaR as a positive number representing potential loss

**Examples:**

```excel
// Calculate 95% VaR
=Risk.VaR(A1:A100, 0.95)

// Calculate 99% VaR
=Risk.VaR(A1:A100, 0.99)
```

**Notes:**

- VaR is returned as a positive number (e.g., 0.02 = 2% loss)
- 95% VaR means: 95% of the time, losses won't exceed this value

---

### Risk.CVaR

**File:** `trading.luau:523`

**Description:** Conditional VaR (Expected Shortfall)

**Category:** Risk Management

**Parameters:**

- `returns` (array): Array of returns
- `confidence_level` (number): Confidence level (default: 0.95)

**Returns:** CVaR (expected loss when VaR is exceeded)

**Examples:**

```excel
// Calculate 95% CVaR
=Risk.CVaR(A1:A100, 0.95)
```

**Notes:**

- CVaR is always >= VaR
- CVaR is the average of losses beyond the VaR threshold
- Also known as Expected Shortfall (ES)

---

## Option Greeks

### Greeks.Delta

**File:** `trading.luau:577`

**Description:** Option delta (sensitivity to underlying price)

**Category:** Option Greeks

**Parameters:**

- `S` (number): Current stock price
- `K` (number): Strike price
- `T` (number): Time to expiration (years)
- `r` (number): Risk-free rate
- `sigma` (number): Volatility
- `option_type` (string): "call" or "put"

**Returns:** Delta value

**Examples:**

```excel
// Calculate delta for a call option
=Greeks.Delta(100, 100, 1, 0.05, 0.20, "call")

// Calculate delta for a put option
=Greeks.Delta(100, 100, 1, 0.05, 0.20, "put")
```

**Notes:**

- Call delta ranges from 0 to 1
- Put delta ranges from -1 to 0
- Delta approximates hedge ratio

---

### Greeks.Gamma

**File:** `trading.luau:599`

**Description:** Option gamma (rate of change of delta)

**Category:** Option Greeks

**Parameters:**

- `S` (number): Current stock price
- `K` (number): Strike price
- `T` (number): Time to expiration (years)
- `r` (number): Risk-free rate
- `sigma` (number): Volatility

**Returns:** Gamma value

**Examples:**

```excel
// Calculate gamma
=Greeks.Gamma(100, 100, 1, 0.05, 0.20)
```

**Notes:**

- Gamma is always positive for both calls and puts
- Highest gamma occurs at-the-money
- Measures convexity of option price

---

### Greeks.Vega

**File:** `trading.luau:613`

**Description:** Option vega (sensitivity to volatility)

**Category:** Option Greeks

**Parameters:**

- `S` (number): Current stock price
- `K` (number): Strike price
- `T` (number): Time to expiration (years)
- `r` (number): Risk-free rate
- `sigma` (number): Volatility

**Returns:** Vega value

**Examples:**

```excel
// Calculate vega
=Greeks.Vega(100, 100, 1, 0.05, 0.20)
```

**Notes:**

- Vega is always positive for both calls and puts
- Represents change in option price for 1% change in volatility
- Highest vega occurs at-the-money

---

### Greeks.Theta

**File:** `trading.luau:627`

**Description:** Option theta (time decay)

**Category:** Option Greeks

**Parameters:**

- `S` (number): Current stock price
- `K` (number): Strike price
- `T` (number): Time to expiration (years)
- `r` (number): Risk-free rate
- `sigma` (number): Volatility
- `option_type` (string): "call" or "put"

**Returns:** Theta value (typically negative)

**Examples:**

```excel
// Calculate theta for a call
=Greeks.Theta(100, 100, 1, 0.05, 0.20, "call")

// Calculate theta for a put
=Greeks.Theta(100, 100, 1, 0.05, 0.20, "put")
```

**Notes:**

- Theta is typically negative (options lose value over time)
- Represents change in option price per day
- Accelerates as expiration approaches

---

## Demo Data Generators

These functions generate realistic sample data for testing and demonstration purposes.

### Demo.PriceSeries

**File:** `trading_demodata.luau:9`

**Description:** Generate sample price series for testing

**Category:** Demo Data

**Parameters:**

- `start_price` (number): Starting price (default: 100)
- `num_days` (number): Number of days to generate (default: 100)
- `annual_return` (number): Expected annual return (default: 0.10)
- `annual_vol` (number): Annual volatility (default: 0.20)
- `seed` (number): Random seed (default: 42)

**Returns:** Array of prices (1 column)

**Examples:**

```excel
// Generate 100 days of price data
=Demo.PriceSeries(100, 100, 0.10, 0.20, 42)
```

---

### Demo.VolumeSeries

**File:** `trading_demodata.luau:51`

**Description:** Generate sample volume data

**Category:** Demo Data

**Parameters:**

- `avg_volume` (number): Average volume (default: 1000000)
- `num_days` (number): Number of days (default: 100)
- `volatility` (number): Volume volatility (default: 0.30)
- `seed` (number): Random seed (default: 123)

**Returns:** Array of volumes (1 column)

**Examples:**

```excel
// Generate volume data
=Demo.VolumeSeries(1000000, 100, 0.30, 123)
```

---

### Demo.ExecutionFills

**File:** `trading_demodata.luau:89`

**Description:** Generate sample execution fill data

**Category:** Demo Data

**Parameters:**

- `target_price` (number): Target execution price (default: 100)
- `total_quantity` (number): Total quantity to fill (default: 10000)
- `num_fills` (number): Number of fills (default: 10)
- `spread_pct` (number): Price spread percentage (default: 0.02 = 2%)
- `seed` (number): Random seed (default: 456)

**Returns:** 2D array with columns: Fill Price, Quantity

**Examples:**

```excel
// Generate execution fills
=Demo.ExecutionFills(100, 10000, 10, 0.02, 456)
```

---

### Demo.OptionChain

**File:** `trading_demodata.luau:134`

**Description:** Generate option chain with strikes around spot

**Category:** Demo Data

**Parameters:**

- `spot_price` (number): Current spot price (default: 100)
- `num_strikes` (number): Number of strikes (default: 11)
- `strike_spacing` (number): Spacing between strikes (default: 5)
- `expiry_years` (number): Years to expiry (default: 1)

**Returns:** Array of strike prices

**Examples:**

```excel
// Generate option chain with strikes from 80 to 120
=Demo.OptionChain(100, 11, 5, 1)
```

---

### Demo.YieldCurve

**File:** `trading_demodata.luau:160`

**Description:** Generate sample yield curve (Nelson-Siegel model)

**Category:** Demo Data

**Parameters:**

- `beta0` (number): Long-term level (default: 0.05)
- `beta1` (number): Short-term component (default: -0.02)
- `beta2` (number): Medium-term component (default: 0.01)
- `lambda` (number): Decay factor (default: 2)

**Returns:** 2D array with columns: Maturity (years), Yield

**Examples:**

```excel
// Generate yield curve
=Demo.YieldCurve(0.05, -0.02, 0.01, 2)
```

**Notes:**

- Uses Nelson-Siegel parameterization
- Maturities: 0.25, 0.5, 1, 2, 3, 5, 7, 10, 20, 30 years

---

### Demo.CorrelationMatrix

**File:** `trading_demodata.luau:188`

**Description:** Generate sample correlation matrix

**Category:** Demo Data

**Parameters:**

- `size` (number): Matrix size (default: 5)
- `avg_correlation` (number): Average correlation (default: 0.3)
- `seed` (number): Random seed (default: 789)

**Returns:** 2D correlation matrix (size × size)

**Examples:**

```excel
// Generate 5×5 correlation matrix
=Demo.CorrelationMatrix(5, 0.3, 789)
```

**Notes:**

- Matrix is symmetric with 1.0 on diagonal
- Correlations bounded between -0.9 and 0.9

---

### Demo.OHLCV

**File:** `trading_demodata.luau:227`

**Description:** Generate OHLCV candlestick data

**Category:** Demo Data

**Parameters:**

- `start_price` (number): Starting price (default: 100)
- `num_candles` (number): Number of candles (default: 50)
- `daily_vol` (number): Daily volatility (default: 0.02)
- `avg_volume` (number): Average volume (default: 1000000)
- `seed` (number): Random seed (default: 999)

**Returns:** 2D array with columns: Open, High, Low, Close, Volume

**Examples:**

```excel
// Generate 50 candles of OHLCV data
=Demo.OHLCV(100, 50, 0.02, 1000000, 999)
```

**Notes:**

- Generates realistic intraday high/low ranges
- Volume is randomized around average

---

### Demo.Portfolio

**File:** `trading_demodata.luau:278`

**Description:** Generate sample portfolio holdings

**Category:** Demo Data

**Parameters:**

- `num_positions` (number): Number of positions (default: 10)
- `total_value` (number): Total portfolio value (default: 1000000)
- `seed` (number): Random seed (default: 321)

**Returns:** 2D array with columns: Ticker, Quantity, Price, Total Value

**Examples:**

```excel
// Generate portfolio with 10 positions
=Demo.Portfolio(10, 1000000, 321)
```

**Notes:**

- Uses major US stock tickers
- Random price allocation between $50-$250

---

## Implementation Notes

### Random Number Generation

Many functions use deterministic random number generators (LCG or PCG algorithms) to ensure reproducibility when the same seed is provided.

### Array Handling

Functions automatically handle both 1D and 2D Excel arrays, flattening 2D arrays when needed.

### Error Handling

Functions return "#ERROR: ..." strings when invalid inputs are provided.

### Generated by xllifyAI

Several functions in `trading.luau` and `trading_demodata.luau` were generated by xllifyAI and are provided for demonstration purposes only. **Do not use in production without thorough testing and validation.**
