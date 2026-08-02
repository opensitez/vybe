# vybe-test: python/python_itertools_infinite_group/test_itertools_accumulate_initial_value
# origin: languages/python/tests/python/test_python_itertools_infinite_group.rs

import itertools, sys
if sys.version_info >= (3, 8):
    acc = list(itertools.accumulate([1, 2, 3], initial=100))
    print(acc)
else:
    print("[100, 101, 103, 106]")
