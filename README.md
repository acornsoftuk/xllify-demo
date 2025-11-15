# xllify-demo

This repo demonstrates using xllify and the [xllify-build](https://github.com/marketplace/actions/xllify-build) action to package some simple Luau functions into an .xll. The workflow defined in [build.yaml](https://github.com/xllifycom/xllify-demo/blob/main/.github/workflows/build.yaml).

For comprehensive documentation of all available functions, see [DOCS.md](./DOCS.md). Note that xllifyAI wrote some of this code so best not to use these for anything other than an illustration, unless of course you know what you're doing.

The action handles the build and packaging process. An .xll file for Microsoft Excel on Windows is attached when a release is created in this repo. See https://xllify.com for more information.

A video demonstrating this repo is here: https://www.youtube.com/watch?v=BvsvYdjh5N8

For more information about the action, see https://github.com/xllifycom/xllify-build or the workflow in this repo.

## Examples

### DuckDB Analytics Demo

The [duckdb](./duckdb) directory contains a comprehensive example of integrating DuckDB's analytical capabilities with Excel. This demo showcases:
- In-memory database with 1.25M rows of synthetic stock market data
- Advanced SQL window functions (moving averages, rankings, percentiles)
- Dynamic array results that spill into Excel cells
- Pre-built analytical queries as Excel functions

See the [DuckDB demo README](./duckdb/README.md) for detailed documentation.

## Testing

If you include a `tests/` directory in your repository, the xllify-build action will automatically discover and run your tests during the build process. If no `tests/` directory exists, the test runner will be skipped and the build will proceed normally. This makes testing completely optional but easy to add when you're ready.

## Installation

After downloading and installing from the [releases page](https://github.com/xllifycom/xllify-demo/releases/latest), Excel should show the shiny new functions.

> Note: you may have to open the .xll file Properties and check the Unblock checkbox.

![Insert function](./screenshots/insert.png)
![Function preview](./screenshots/preview.png)
![All](./screenshots/all.png)
