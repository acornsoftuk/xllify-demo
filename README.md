# xllify-demo

This repo demonstrates using xllify and the [xllify-build](https://github.com/marketplace/actions/xllify-build) action to package some simple Luau functions into an .xll. The workflows are defined in [build.yaml](https://github.com/xllifycom/xllify-demo/blob/main/.github/workflows/build.yaml).

You can download pre-built XLLs from the [latest release](https://github.com/xllifycom/xllify-demo/releases/latest).

For more information about the action, see https://github.com/xllifycom/xllify-build or the workflow in this repo.

## Further examples

### DuckDB

The [duckdb](./duckdb) directory contains a comprehensive example of integrating DuckDB's analytical capabilities with Excel. This demo showcases:
- In-memory database with 1.25M rows of synthetic stock market data
- Advanced SQL window functions (moving averages, rankings, percentiles)
- Dynamic array results that spill into Excel cells
- Pre-built analytical queries as Excel functions

See the [DuckDB demo README](./duckdb/README.md) for detailed documentation. It is built as xllify_demo_duckdb.xll and is attached to releases.

## Installation

After downloading and installing from the [releases page](https://github.com/xllifycom/xllify-demo/releases/latest), Excel should show the shiny new functions.

> Note: you may have to open the .xll file Properties and check the Unblock checkbox.

![Insert function](./screenshots/insert.png)
![Function preview](./screenshots/preview.png)
![All](./screenshots/all.png)
