# vybe-test: python/bisect_heapq_stdlib/statistics_linear_regression
# origin: languages/python/tests/python/test_bisect_heapq_stdlib.rs
# `linear_regression(x, y)` takes TWO sequences, not a list of pairs.
import statistics
statistics.linear_regression([1, 2], [2, 4])
