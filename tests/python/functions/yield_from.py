# vybe-test: python/functions/yield_from
# origin: languages/python/tests/python/test_functions.rs

def chain(a, b):
    yield from a
    yield from b
