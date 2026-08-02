# vybe-test: python/random_stdlib/random_sample_k_too_large
# origin: languages/python/tests/python/test_random_stdlib.rs
# vybe-test-mode: compile

import random
try:
    random.sample([1, 2], 3)
except ValueError:
    pass
