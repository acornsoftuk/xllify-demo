# perftest

Extremely simple performance test to exercise calling a function 5000 times. The function implementation is just a simple random number generator. Download a pre-built version [from the latest release](https://github.com/xllifycom/xllify-demo/releases/latest)

The sheet `exercise.xlsm` contains 5000 rows. Hit one of the buttons to recalculate. This workbook will work for both the luau and Python variants.

- Load exercise.xlsx
- Load xllify_demo_luau.xll or xllify_demo_python.xll
- Observe recalc
- Change seed value
- Observe recalc

Luau uses multithreaded recalculation. Python starts 3 worker processes (configurable with `configure_spawn_count`) and load balances between them. This avoids a single core being monopolised or the GIL slowing things down.

