# vybe-test: python/functions/yield_from
# origin: languages/python/tests/python/test_functions.rs
# vybe-test-mode: compile

def chain(a, b):
    yield from a
    yield from b
