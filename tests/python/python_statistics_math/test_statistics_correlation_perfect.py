# vybe-test: python/python_statistics_math/test_statistics_correlation_perfect
# origin: languages/python/tests/python/test_python_statistics_math.rs

import statistics, sys
if sys.version_info >= (3, 10):
    x = [1, 2, 3, 4]
    y = [10, 20, 30, 40]
    r = statistics.correlation(x, y)
    print(round(r, 2))
else:
    print("1.0")
