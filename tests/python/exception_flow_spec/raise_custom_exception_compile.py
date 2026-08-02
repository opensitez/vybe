# vybe-test: python/exception_flow_spec/raise_custom_exception_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

class MyError(Exception):
    pass
raise MyError('boom')
