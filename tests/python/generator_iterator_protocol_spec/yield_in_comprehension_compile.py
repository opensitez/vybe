# vybe-test: python/generator_iterator_protocol_spec/yield_in_comprehension_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs

def gen():
    for x in [1, 2, 3]:
        yield x * 2
