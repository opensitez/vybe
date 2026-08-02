# vybe-test: python/generator_iterator_protocol_spec/iter_default_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

it = iter([1, 2, 3])
x = next(it, None)
