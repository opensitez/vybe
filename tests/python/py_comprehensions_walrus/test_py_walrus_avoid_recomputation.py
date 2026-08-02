# vybe-test: python/py_comprehensions_walrus/test_py_walrus_avoid_recomputation
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

import math

data = [4, 9, 16, -1, 25]
# compute once in condition, reuse in body
results = []
for n in data:
    if (sq := math.sqrt(n)) > 3 if n >= 0 else False:
        results.append(sq)
print(results)
