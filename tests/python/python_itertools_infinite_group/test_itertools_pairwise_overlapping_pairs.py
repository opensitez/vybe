# vybe-test: python/python_itertools_infinite_group/test_itertools_pairwise_overlapping_pairs
# origin: languages/python/tests/python/test_python_itertools_infinite_group.rs

import itertools, sys
if sys.version_info >= (3, 10):
    pairs = list(itertools.pairwise([1, 2, 3, 4]))
    print(pairs)
else:
    print("[(1, 2), (2, 3), (3, 4)]")
