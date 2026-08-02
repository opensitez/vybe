# vybe-test: python/generator_iterator_protocol_spec/iter_callable_sentinel_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

def reader():
    return 0
it = iter(reader, 0)
