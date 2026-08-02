# vybe-test: python/py_itertools_infinite_combinatorics/test_py_itertools_groupby_consecutive_runs
# origin: languages/python/tests/python/test_py_itertools_infinite_combinatorics.rs

from itertools import groupby

data = [1, 1, 2, 3, 3, 3, 2, 2, 1]
grouped = [(k, list(g)) for k, g in groupby(data)]
print(grouped)
