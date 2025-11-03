# xllify-demo Function Documentation

This document provides comprehensive documentation for all custom Excel functions implemented in this repository using [xllify.com](https://xllify.com).

> **Note:** Some examples in this repository were generated with xllifyAI and may contain errors. Always validate results before using in production environments.

## Table of Contents

- [Simple Demo Functions](#simple-demo-functions)
- [Demo Data Generators](#demo-data-generators)
- [Black-Scholes Option Pricing](#black-scholes-option-pricing)
- [Portfolio Analytics](#portfolio-analytics)
- [Market Simulation](#market-simulation)
- [Trading Analytics](#trading-analytics)
- [Technical Analysis](#technical-analysis)
- [Fixed Income](#fixed-income)
- [Risk Management](#risk-management)
- [Option Greeks](#option-greeks)

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

## Demo Data Generators

These functions generate sample data for testing and demonstration purposes.

### Demo.PriceSeries

**File:** `trading_demodata.luau:9`

**Description:** Generate sample price series for testing using geometric Brownian motion.

**Category:** Demo Data

**Parameters:**

- `start_price` (number): Starting price (default: 100)
- `num_days` (number): Number of days to generate (default: 100)
- `annual_return` (number): Expected annual return (default: 0.10)
- `annual_vol` (number): Annual volatility (default: 0.20)
- `seed` (number): Random seed for reproducibility (default: 42)

**Returns:** Array of prices (single column)

**Examples:**

```excel
// Generate 100 days of prices starting at $100
=Demo.PriceSeries(100, 100, 0.10, 0.20, 42)

// Generate 252 days with higher volatility
=Demo.PriceSeries(50, 252, 0.08, 0.35)
```

---

### Demo.VolumeSeries

**File:** `trading_demodata.luau:51`

**Description:** Generate sample volume data with realistic variation.

**Category:** Demo Data

**Parameters:**

- `avg_volume` (number): Average volume (default: 1,000,000)
- `num_days` (number): Number of days (default: 100)
- `volatility` (number): Volume volatility (default: 0.30)
- `seed` (number): Random seed (default: 123)

**Returns:** Array of volume values (single column)

**Examples:**

```excel
// Generate 100 days of volume data
=Demo.VolumeSeries(1000000, 100, 0.30, 123)
```

---

### Demo.ExecutionFills

**File:** `trading_demodata.luau:89`

**Description:** Generate sample execution fill data showing multiple fills at different prices.

**Category:** Demo Data

**Parameters:**

- `target_price` (number): Target execution price (default: 100)
- `total_quantity` (number): Total quantity to fill (default: 10,000)
- `num_fills` (number): Number of separate fills (default: 10)
- `spread_pct` (number): Price spread percentage (default: 0.02)
- `seed` (number): Random seed (default: 456)

**Returns:** 2D array with columns: Fill Price, Fill Quantity

**Examples:**

```excel
// Generate 10 fills for 10,000 shares around $100
=Demo.ExecutionFills(100, 10000, 10, 0.02, 456)
```

---

### Demo.OptionChain

**File:** `trading_demodata.luau:134`

**Description:** Generate option chain with strike prices around the spot price.

**Category:** Demo Data

**Parameters:**

- `spot_price` (number): Current stock price (default: 100)
- `num_strikes` (number): Number of strikes to generate (default: 11)
- `strike_spacing` (number): Spacing between strikes (default: 5)
- `expiry_years` (number): Years to expiration (default: 1)

**Returns:** Array of strike prices (single column)

**Examples:**

```excel
// Generate 11 strikes around $100 with $5 spacing
=Demo.OptionChain(100, 11, 5, 1)
```

---

### Demo.YieldCurve

**File:** `trading_demodata.luau:160`

**Description:** Generate sample yield curve using the Nelson-Siegel model.

**Category:** Demo Data

**Parameters:**

- `beta0` (number): Long-term level (default: 0.05)
- `beta1` (number): Short-term component (default: -0.02)
- `beta2` (number): Medium-term component (default: 0.01)
- `lambda` (number): Decay factor (default: 2)

**Returns:** 2D array with columns: Maturity (years), Yield

**Examples:**

```excel
// Generate yield curve with default parameters
=Demo.YieldCurve(0.05, -0.02, 0.01, 2)
```

---

### Demo.CorrelationMatrix

**File:** `trading_demodata.luau:188`

**Description:** Generate sample correlation matrix.

**Category:** Demo Data

**Parameters:**

- `size` (number): Matrix dimension (default: 5)
- `avg_correlation` (number): Average correlation (default: 0.3)
- `seed` (number): Random seed (default: 789)

**Returns:** Square correlation matrix

**Examples:**

```excel
// Generate 5x5 correlation matrix
=Demo.CorrelationMatrix(5, 0.3, 789)
```

---

### Demo.OHLCV

**File:** `trading_demodata.luau:227`

**Description:** Generate OHLCV (Open, High, Low, Close, Volume) candlestick data.

**Category:** Demo Data

**Parameters:**

- `start_price` (number): Starting price (default: 100)
- `num_candles` (number): Number of candles (default: 50)
- `daily_vol` (number): Daily volatility (default: 0.02)
- `avg_volume` (number): Average volume (default: 1,000,000)
- `seed` (number): Random seed (default: 999)

**Returns:** 2D array with columns: Open, High, Low, Close, Volume

**Examples:**

```excel
// Generate 50 candles starting at $100
=Demo.OHLCV(100, 50, 0.02, 1000000, 999)
```

---

### Demo.Portfolio

**File:** `trading_demodata.luau:278`

**Description:** Generate sample portfolio holdings.

**Category:** Demo Data

**Parameters:**

- `num_positions` (number): Number of positions (default: 10)
- `total_value` (number): Total portfolio value (default: 1,000,000)
- `seed` (number): Random seed (default: 321)

**Returns:** 2D array with columns: Ticker, Quantity, Price, Total Value

**Examples:**

```excel
// Generate 10-position portfolio worth $1M
=Demo.Portfolio(10, 1000000, 321)
```

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

### bs.State

**File:** `black_scholes.luau:103`

**Description:** Determine market state from Black-Scholes call and put prices

**Parameters:**

- `callPrice` (number): Call option price
- `putPrice` (number): Put option price
- `S` (number): Current stock price
- `K` (number): Strike price
- `T` (number): Time to expiration (years)
- `r` (number): Risk-free interest rate

**Returns:** String describing the market state

**Examples:**

```excel
// Analyze market state from option prices
=bs.State(10.45, 5.57, 100, 100, 1, 0.05)
// Returns: "ATM" (at-the-money)

// Check deep in-the-money scenario
=bs.State(15.00, 2.50, 110, 100, 1, 0.05)
// Returns: "Deep ITM Call / OTM Put"

// Detect put-call parity violations
=bs.State(12.00, 8.00, 100, 100, 1, 0.05)
// May return: "ATM [Parity Violation]"
```

**Market States:**

- **Deep ITM Call / OTM Put**: Spot > 1.05 × Strike
- **ITM Call / OTM Put**: Spot > 1.01 × Strike
- **ATM**: 0.99 ≤ Spot/Strike ≤ 1.01
- **OTM Call / ITM Put**: 0.95 ≤ Spot/Strike < 0.99
- **OTM Call / Deep ITM Put**: Spot < 0.95 × Strike

**Notes:**

- Uses put-call parity to detect arbitrage opportunities: C - P = S - K×e^(-rT)
- Flags "[Parity Violation]" if deviation exceeds 1% of spot price
- Useful for validating option pricing consistency
- Moneyness categories based on spot/strike ratio

---

## Portfolio Analytics

Functions for analyzing portfolio performance and risk metrics.

### Portfolio.Returns

**File:** `trading.luau:10`

**Description:** Calculate returns from a price series.

**Category:** Portfolio Analytics

**Parameters:**

- `prices` (array): Array of prices

**Returns:** Array of returns (one fewer element than input)

**Examples:**

```excel
// Calculate returns from price data in A1:A100
=Portfolio.Returns(A1:A100)
```

**Notes:**

- Returns are calculated as simple returns: (price[i] - price[i-1]) / price[i-1]
- Handles both 1D and 2D arrays

---

### Portfolio.Volatility

**File:** `trading.luau:35`

**Description:** Calculate annualized volatility from returns.

**Category:** Portfolio Analytics

**Parameters:**

- `returns` (array): Array of returns
- `periods_per_year` (number): Number of periods per year (default: 252 for daily data)

**Returns:** Annualized volatility

**Examples:**

```excel
// Calculate volatility from daily returns
=Portfolio.Volatility(A1:A100, 252)

// Calculate volatility from monthly returns
=Portfolio.Volatility(A1:A100, 12)
```

**Notes:**

- Uses sample standard deviation (n-1 denominator)
- Annualization assumes independent returns

---

### Portfolio.Sharpe

**File:** `trading.luau:78`

**Description:** Calculate Sharpe ratio (excess return divided by volatility).

**Category:** Portfolio Analytics

**Parameters:**

- `returns` (array): Array of returns
- `risk_free_rate` (number): Annual risk-free rate (default: 0.02)
- `periods_per_year` (number): Periods per year (default: 252)

**Returns:** Sharpe ratio

**Examples:**

```excel
// Calculate Sharpe ratio with 2% risk-free rate
=Portfolio.Sharpe(A1:A100, 0.02, 252)
```

**Notes:**

- Higher Sharpe ratios indicate better risk-adjusted returns
- Uses annualized values

---

## Market Simulation

Functions for simulating market scenarios and pricing via Monte Carlo.

### Simulation.PricePath

**File:** `trading.luau:139`

**Description:** Simulate geometric Brownian motion price path.

**Category:** Market Simulation

**Parameters:**

- `S0` (number): Initial stock price (default: 100)
- `mu` (number): Expected annual return (default: 0.10)
- `sigma` (number): Annual volatility (default: 0.20)
- `days` (number): Number of days to simulate (default: 252)
- `seed` (number): Random seed (default: 42)

**Returns:** Array of simulated prices

**Examples:**

```excel
// Simulate 1 year of daily prices starting at $100
=Simulation.PricePath(100, 0.10, 0.20, 252, 42)
```

**Notes:**

- Uses geometric Brownian motion: dS = μS dt + σS dW
- Box-Muller transform for normal random variables

---

### Simulation.MonteCarlo

**File:** `trading.luau:165`

**Description:** Run Monte Carlo option pricing simulation.

**Category:** Market Simulation

**Parameters:**

- `S` (number): Current stock price (default: 100)
- `K` (number): Strike price (default: 100)
- `T` (number): Time to expiration in years (default: 1)
- `r` (number): Risk-free rate (default: 0.05)
- `sigma` (number): Volatility (default: 0.20)
- `num_sims` (number): Number of simulations (default: 10,000)
- `option_type` (string): "call" or "put" (default: "call")

**Returns:** Option price estimate

**Examples:**

```excel
// Price a call option with 100,000 simulations
=Simulation.MonteCarlo(100, 100, 1, 0.05, 0.20, 100000, "call")

// Price a put option
=Simulation.MonteCarlo(100, 100, 1, 0.05, 0.20, 10000, "put")
```

**Notes:**

- More simulations increase accuracy but take longer to compute
- Results will vary slightly between calls due to randomness

---

## Trading Analytics

Functions for analyzing trade execution quality.

### Trading.VWAP

**File:** `trading.luau:204`

**Description:** Calculate volume-weighted average price.

**Category:** Trading Analytics

**Parameters:**

- `prices` (array): Array of prices
- `volumes` (array): Array of volumes

**Returns:** VWAP value

**Examples:**

```excel
// Calculate VWAP from price and volume data
=Trading.VWAP(A1:A100, B1:B100)
```

**Notes:**

- VWAP = Σ(price × volume) / Σ(volume)
- Commonly used as execution benchmark

---

### Trading.Slippage

**File:** `trading.luau:235`

**Description:** Calculate execution slippage versus benchmark price.

**Category:** Trading Analytics

**Parameters:**

- `benchmark_price` (number): Benchmark price (e.g., arrival price)
- `execution_prices` (array): Array of execution prices
- `quantities` (array): Array of execution quantities

**Returns:** Slippage as percentage of benchmark

**Examples:**

```excel
// Calculate slippage vs $100 benchmark
=Trading.Slippage(100, A1:A10, B1:B10)
```

**Notes:**

- Positive slippage means execution was worse than benchmark
- Also known as implementation shortfall

---

## Technical Analysis

Classic technical indicators for price analysis.

### Technical.SMA

**File:** `trading.luau:270`

**Description:** Simple moving average.

**Category:** Technical Analysis

**Parameters:**

- `prices` (array): Price series
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

- Output array has fewer elements than input (length - window + 1)
- Each value is the average of the previous `window` prices

---

### Technical.EMA

**File:** `trading.luau:306`

**Description:** Exponential moving average.

**Category:** Technical Analysis

**Parameters:**

- `prices` (array): Price series
- `window` (number): Window size (default: 20)

**Returns:** Array of EMA values

**Examples:**

```excel
// Calculate 20-period EMA
=Technical.EMA(A1:A100, 20)
```

**Notes:**

- EMA gives more weight to recent prices
- Smoothing factor α = 2/(window+1)
- First EMA value is calculated as SMA

---

### Technical.RSI

**File:** `trading.luau:349`

**Description:** Relative Strength Index.

**Category:** Technical Analysis

**Parameters:**

- `prices` (array): Price series
- `period` (number): RSI period (default: 14)

**Returns:** Array of RSI values (0-100)

**Examples:**

```excel
// Calculate 14-period RSI
=Technical.RSI(A1:A100, 14)
```

**Notes:**

- RSI ranges from 0 to 100
- Traditionally, >70 indicates overbought, <30 indicates oversold
- Uses smoothed average gains and losses

---

## Fixed Income

Bond pricing and duration calculations.

### Bond.Price

**File:** `trading.luau:429`

**Description:** Calculate bond price from yield.

**Category:** Fixed Income

**Parameters:**

- `face_value` (number): Par value (default: 1000)
- `coupon_rate` (number): Annual coupon rate (default: 0.05)
- `ytm` (number): Yield to maturity (default: 0.06)
- `years` (number): Years to maturity (default: 10)
- `frequency` (number): Coupon frequency per year (default: 2 for semi-annual)

**Returns:** Bond price

**Examples:**

```excel
// Price a 10-year bond with 5% coupon, 6% YTM
=Bond.Price(1000, 0.05, 0.06, 10, 2)
```

**Notes:**

- Price below par when YTM > coupon rate
- Price above par when YTM < coupon rate

---

### Bond.Duration

**File:** `trading.luau:457`

**Description:** Calculate Macaulay duration.

**Category:** Fixed Income

**Parameters:**

- `face_value` (number): Par value (default: 1000)
- `coupon_rate` (number): Annual coupon rate (default: 0.05)
- `ytm` (number): Yield to maturity (default: 0.06)
- `years` (number): Years to maturity (default: 10)
- `frequency` (number): Coupon frequency (default: 2)

**Returns:** Duration in years

**Examples:**

```excel
// Calculate duration of 10-year bond
=Bond.Duration(1000, 0.05, 0.06, 10, 2)
```

**Notes:**

- Measures weighted average time to receive cash flows
- Used to estimate price sensitivity to yield changes

---

## Risk Management

Value at Risk and Expected Shortfall calculations.

### Risk.VaR

**File:** `trading.luau:494`

**Description:** Calculate Value at Risk using historical method.

**Category:** Risk Management

**Parameters:**

- `returns` (array): Historical returns
- `confidence_level` (number): Confidence level (default: 0.95)

**Returns:** VaR (as positive number for losses)

**Examples:**

```excel
// Calculate 95% VaR
=Risk.VaR(A1:A100, 0.95)

// Calculate 99% VaR
=Risk.VaR(A1:A100, 0.99)
```

**Notes:**

- VaR represents the maximum expected loss at given confidence level
- Uses historical simulation (no distribution assumption)
- Result is returned as positive number

---

### Risk.CVaR

**File:** `trading.luau:522`

**Description:** Conditional VaR (Expected Shortfall) - average loss beyond VaR threshold.

**Category:** Risk Management

**Parameters:**

- `returns` (array): Historical returns
- `confidence_level` (number): Confidence level (default: 0.95)

**Returns:** CVaR (as positive number)

**Examples:**

```excel
// Calculate 95% CVaR (Expected Shortfall)
=Risk.CVaR(A1:A100, 0.95)
```

**Notes:**

- CVaR is always >= VaR
- Provides estimate of tail risk beyond VaR
- Also known as Expected Shortfall (ES) or Average VaR

---

## Option Greeks

Sensitivity measures for options (delta, gamma, vega, theta).

### Greeks.Delta

**File:** `trading.luau:576`

**Description:** Option delta - sensitivity to underlying price changes.

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
// Calculate delta for call option
=Greeks.Delta(100, 100, 1, 0.05, 0.20, "call")

// Calculate delta for put option
=Greeks.Delta(100, 100, 1, 0.05, 0.20, "put")
```

**Notes:**

- Call delta ranges from 0 to 1
- Put delta ranges from -1 to 0
- Represents hedge ratio (shares of stock per option)

---

### Greeks.Gamma

**File:** `trading.luau:598`

**Description:** Option gamma - rate of change of delta.

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
// Calculate gamma (same for calls and puts)
=Greeks.Gamma(100, 100, 1, 0.05, 0.20)
```

**Notes:**

- Gamma is highest for at-the-money options
- Same value for calls and puts
- Measures convexity of option price

---

### Greeks.Vega

**File:** `trading.luau:612`

**Description:** Option vega - sensitivity to volatility changes.

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
// Calculate vega (same for calls and puts)
=Greeks.Vega(100, 100, 1, 0.05, 0.20)
```

**Notes:**

- Vega is highest for at-the-money options
- Same value for calls and puts
- Typically quoted per 1% change in volatility

---

### Greeks.Theta

**File:** `trading.luau:626`

**Description:** Option theta - time decay.

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
// Calculate theta for call option
=Greeks.Theta(100, 100, 1, 0.05, 0.20, "call")

// Calculate theta for put option
=Greeks.Theta(100, 100, 1, 0.05, 0.20, "put")
```

**Notes:**

- Typically negative for long positions (options lose value over time)
- Accelerates as expiration approaches
- Usually quoted per day (divide by 365 for daily theta)
