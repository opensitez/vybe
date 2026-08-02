# vybe-test: python/exception_flow_spec/reraise_bare_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

try:
    risky()
except Exception:
    raise
