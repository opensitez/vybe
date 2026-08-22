# vybe-test: python/random_stdlib/random_randrange_zero_step_raises
# origin: languages/python/tests/python/test_random_stdlib.rs

import random
try:
    random.randrange(0, 10, 0)
except ValueError:
    pass
