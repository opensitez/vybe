# vybe-test: python/functools_itertools/itertools_groupby
# origin: languages/python/tests/python/test_functools_itertools.rs

import itertools
data = [('a', 1), ('a', 2), ('b', 3)]
print([k for k, _ in itertools.groupby(data, key=lambda x: x[0])])
