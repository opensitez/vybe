# vybe-test: python/slicing_unpacking_spec/unpack_generator_compile
# origin: languages/python/tests/python/test_slicing_unpacking_spec.rs
# vybe-test-mode: compile

def gen():
    yield 1
    yield 2
a, b = gen()
