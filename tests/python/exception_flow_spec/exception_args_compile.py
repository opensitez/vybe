# vybe-test: python/exception_flow_spec/exception_args_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

exc = ValueError('bad', 3)
args = exc.args
