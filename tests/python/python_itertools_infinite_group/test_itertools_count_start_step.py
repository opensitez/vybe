# vybe-test: python/python_itertools_infinite_group/test_itertools_count_start_step
# origin: languages/python/tests/python/test_python_itertools_infinite_group.rs

import itertools
counter = itertools.count(start=10, step=2.5)
vals = [next(counter) for _ in range(4)]
print(vals)
