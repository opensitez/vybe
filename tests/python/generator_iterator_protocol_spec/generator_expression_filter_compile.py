# vybe-test: python/generator_iterator_protocol_spec/generator_expression_filter_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs

x = (i for i in range(10) if i % 2 == 0)
