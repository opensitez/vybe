# vybe-test: python/py_generator_coroutines_yield/test_py_yield_from_iterable_delegation
# origin: languages/python/tests/python/test_py_generator_coroutines_yield.rs

def delegate_all():
    yield from [1, 2, 3]
    yield from (x * 10 for x in range(1, 4))
    yield from "AB"

print(list(delegate_all()))
