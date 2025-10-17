# xllify-demo Function Documentation

This document provides comprehensive documentation for all custom Excel functions implemented in this repository using [xllify.com](https://xllify.com).

## Table of Contents

- [Simple Demo Functions](#simple-demo-functions)
- [Black-Scholes Option Pricing](#black-scholes-option-pricing)

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
