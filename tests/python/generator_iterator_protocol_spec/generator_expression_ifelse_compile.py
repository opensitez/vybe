# vybe-test: python/generator_iterator_protocol_spec/generator_expression_ifelse_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

x = ('even' if i % 2 == 0 else 'odd' for i in range(4))
