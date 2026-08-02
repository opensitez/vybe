# vybe-test: python/py_operator/test_py_operator_in_functional_pipeline
# origin: languages/python/tests/python/test_py_operator.rs

import operator
from functools import reduce

numbers = [1, 2, 3, 4, 5]
total = reduce(operator.add, numbers)
product = reduce(operator.mul, numbers)
print(total)
print(product)

# Build matrix of comparisons
pairs = [(1, 2), (5, 3), (4, 4)]
results = [operator.lt(a, b) for a, b in pairs]
print(results)
