# vybe-test: python/py_random_secrets_prng/test_py_random_seed_reproducibility
# origin: languages/python/tests/python/test_py_random_secrets_prng.rs

import random

random.seed(42)
r1 = [random.randint(1, 100) for _ in range(5)]

random.seed(42)
r2 = [random.randint(1, 100) for _ in range(5)]

print(r1 == r2)
print(r1)
