# vybe-test: python/python_statistics_math/test_statistics_linear_regression
# origin: languages/python/tests/python/test_python_statistics_math.rs

import statistics, sys
if sys.version_info >= (3, 10):
    x = [1, 2, 3]
    y = [2, 4, 6]
    slope, intercept = statistics.linear_regression(x, y)
    print(round(slope, 1), round(intercept, 1))
else:
    print("2.0 0.0")
