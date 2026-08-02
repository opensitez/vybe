# vybe-test: python/py_generators_iterators/test_py_generator_infinite_sequence
# origin: languages/python/tests/python/test_py_generators_iterators.rs

def fibonacci():
    a, b = 0, 1
    while True:
        yield a
        a, b = b, a + b

import itertools
print(list(itertools.islice(fibonacci(), 8)))
