# vybe-test: python/generator_iterator_protocol_spec/yield_from_list_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

def gen():
    yield from [1, 2, 3]
