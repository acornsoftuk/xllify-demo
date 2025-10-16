# xllify-demo

This repo demonstrates using [xllify.com](https://xllify.com) by implementing some simple custom functions in Luau. The xllify service handles the build and packaging process. An .xll file for Windows Excel is attached when a release is created in this repo.

A good way to get started is to [fork this repo](https://github.com/acornsoftuk/xllify-demo/fork) (remember to enable actions in your fork.)

For more information about the action, see https://github.com/acornsoftuk/xllify-build or the workflow in this repo.

After downloading and installing from the [releases page](https://github.com/acornsoftuk/xllify-demo/releases/latest), Excel should show these shiny new functions.

![Insert function](./screenshots/insert.png)
![Function preview](./screenshots/preview.png)
![All](./screenshots/all.png)

## Available functions

These are all defined in [hello.luau](./hello.luau). Note that they are intended only as illustrations to demonstrate xllify, their logic may not be entirely correct.

### xllify.Demo.Hello

**Description:** Says hello!

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

**Description:** Determines your life stage based on age

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

**Category:** xllify Demos

**Description:** Generate sample portfolio holdings

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
