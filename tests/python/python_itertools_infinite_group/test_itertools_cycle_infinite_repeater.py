# vybe-test: python/python_itertools_infinite_group/test_itertools_cycle_infinite_repeater
# origin: languages/python/tests/python/test_python_itertools_infinite_group.rs

import itertools
cycler = itertools.cycle(["red", "green", "blue"])
vals = [next(cycler) for _ in range(5)]
print(vals)
