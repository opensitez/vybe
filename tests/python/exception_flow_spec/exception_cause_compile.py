# vybe-test: python/exception_flow_spec/exception_cause_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

try:
    raise ValueError()
except ValueError as exc:
    raise RuntimeError() from exc
