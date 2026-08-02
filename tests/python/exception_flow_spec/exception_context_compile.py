# vybe-test: python/exception_flow_spec/exception_context_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

try:
    raise ValueError()
except ValueError as exc:
    ctx = exc.__context__
