# vybe-test: python/exception_flow_spec/assert_raises_custom_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

class MyError(Exception):
    pass
try:
    raise MyError()
except MyError:
    pass
