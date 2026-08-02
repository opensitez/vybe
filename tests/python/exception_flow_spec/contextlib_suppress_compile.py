# vybe-test: python/exception_flow_spec/contextlib_suppress_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

from contextlib import suppress
with suppress(ValueError):
    raise ValueError()
