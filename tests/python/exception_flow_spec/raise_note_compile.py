# vybe-test: python/exception_flow_spec/raise_note_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

exc = ValueError('bad')
exc.add_note('context')
raise exc
