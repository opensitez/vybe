# vybe-test: python/generator_iterator_protocol_spec/generator_return_value_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

def gen():
    yield 1
    return 99
