# vybe-test: python/exception_flow_spec/reraise_bare_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
def risky(*_a, **_k):
    return None

try:
    risky()
except Exception:
    raise
