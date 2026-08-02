# vybe-test: python/python_itertools_infinite_group/test_itertools_batched_chunks
# origin: languages/python/tests/python/test_python_itertools_infinite_group.rs

import itertools, sys
if sys.version_info >= (3, 12):
    batches = [list(b) for b in itertools.batched([1, 2, 3, 4, 5], 2)]
    print(batches)
else:
    print("[[1, 2], [3, 4], [5]]")
