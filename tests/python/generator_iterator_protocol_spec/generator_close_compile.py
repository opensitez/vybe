# vybe-test: python/generator_iterator_protocol_spec/generator_close_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs

def gen():
    yield 1

g = gen()
g.close()
