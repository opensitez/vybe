# vybe-test: python/py_itertools/test_py_itertools_groupby_consecutive
# origin: languages/python/tests/python/test_py_itertools.rs

import itertools

data = [("A", 1), ("A", 2), ("B", 3), ("A", 4)]
groups = {k: list(v) for k, v in itertools.groupby(data, key=lambda x: x[0])}
print(groups)
