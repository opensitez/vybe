# vybe-test: python/exception_flow_spec/nested_finally_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
def risky(*_a, **_k):
    return None
def clean_inner(*_a, **_k):
    return None
def clean_outer(*_a, **_k):
    return None

try:
    try:
        risky()
    finally:
        clean_inner()
finally:
    clean_outer()
