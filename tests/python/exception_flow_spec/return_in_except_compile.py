# vybe-test: python/exception_flow_spec/return_in_except_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs

def f():
    try:
        risky()
    except Exception:
        return 1
