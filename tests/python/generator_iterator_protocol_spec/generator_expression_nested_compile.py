# vybe-test: python/generator_iterator_protocol_spec/generator_expression_nested_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

x = ((i, j) for i in range(2) for j in range(2))
