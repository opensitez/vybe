# vybe-test: python/random_stdlib/random_choices_cum_weights
# origin: languages/python/tests/python/test_random_stdlib.rs
# vybe-test-mode: compile

import random
random.choices([1, 2], cum_weights=[1, 3], k=2)
