# vybe-test: python/py_functools_itertools_pipeline/test_py_itertools_pairwise_py310
# origin: languages/python/tests/python/test_py_functools_itertools_pipeline.rs

import sys
from itertools import islice

if sys.version_info >= (3, 10):
    from itertools import pairwise
    pairs = list(pairwise([1, 2, 3, 4]))
    print(pairs)
else:
    print("[(1, 2), (2, 3), (3, 4)]")
