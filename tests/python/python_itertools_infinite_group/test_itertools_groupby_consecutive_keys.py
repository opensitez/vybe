# vybe-test: python/python_itertools_infinite_group/test_itertools_groupby_consecutive_keys
# origin: languages/python/tests/python/test_python_itertools_infinite_group.rs

import itertools
data = [("a", 1), ("a", 2), ("b", 3), ("b", 4), ("a", 5)]
res = []
for k, g in itertools.groupby(data, key=lambda x: x[0]):
    res.append((k, [item[1] for item in g]))
print(res)
