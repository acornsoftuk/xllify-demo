# perftest

Extremely simple performance test to exercise calling a function 5000 times. The function implementation is just a simple random number generator.

The sheet `exercise.xlsx` contains 5000 rows. Change the seed cell to recalculate. This workbook will work for both the luau and Python variants, but not at the same time of course.

- Load exercise.xlsx
- Load xllify_demo_luau.xll or xllify_demo_python.xll
- Observe recalc
- Change seed value
- Observe recalc

Luau uses multithreaded recalculation. Multi-process support for Python is coming soon. (It's actually in the codebase already just not exposed.)
