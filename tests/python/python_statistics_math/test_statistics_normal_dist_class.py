# vybe-test: python/python_statistics_math/test_statistics_normal_dist_class
# origin: languages/python/tests/python/test_python_statistics_math.rs

import statistics, sys
if sys.version_info >= (3, 8):
    nd = statistics.NormalDist(mu=10, sigma=2)
    print(nd.mean)
    print(nd.stdev)
else:
    print("10.0")
    print("2.0")
