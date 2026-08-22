# vybe-test: python/generator_extended/generator_yield_from_generator
# origin: languages/python/tests/python/test_generator_extended.rs

def inner():
 yield 1
def outer():
 yield from inner()
list(outer())
