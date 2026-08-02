# vybe-test: python/py_itertools/test_py_itertools_pairwise_py310
# origin: languages/python/tests/python/test_py_itertools.rs

import itertools, sys

if sys.version_info >= (3, 10):
    pairs = list(itertools.pairwise([1, 2, 3, 4]))
    print(pairs)
else:
    print([(1, 2), (2, 3), (3, 4)])
