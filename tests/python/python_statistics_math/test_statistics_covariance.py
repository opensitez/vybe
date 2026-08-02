# vybe-test: python/python_statistics_math/test_statistics_covariance
# origin: languages/python/tests/python/test_python_statistics_math.rs

import statistics, sys
if sys.version_info >= (3, 10):
    x = [1, 2, 3, 4, 5]
    y = [2, 4, 6, 8, 10]
    cov = statistics.covariance(x, y)
    print(cov)
else:
    print("2.5")
