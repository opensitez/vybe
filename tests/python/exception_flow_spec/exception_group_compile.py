# vybe-test: python/exception_flow_spec/exception_group_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

group = ExceptionGroup('many', [ValueError('a'), TypeError('b')])
