# vybe-test: python/exception_flow_spec/finally_return_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs

def f():
    try:
        return 1
    finally:
        return 2
