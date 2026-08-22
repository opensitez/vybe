# vybe-test: python/exception_flow_spec/except_baseexception_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs

try:
    risky()
except BaseException:
    pass
