# vybe-test: python/random_stdlib/random_randint_swapped_raises
# origin: languages/python/tests/python/test_random_stdlib.rs

import random
try:
    random.randint(10, 5)
    print('ok')
except ValueError:
    print('bad')
