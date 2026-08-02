# vybe-test: python/python_itertools_grouping/test_python_itertools_groupby_keyed
# origin: languages/python/tests/python/test_python_itertools_grouping.rs

import itertools
items = [('a', 1), ('a', 2), ('b', 1), ('b', 3)]
for key, group in itertools.groupby(items, lambda x: x[0]):
    vals = [v[1] for v in group]
    print(f"{key}:{sum(vals)}")
