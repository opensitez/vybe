# vybe-test: python/context_managers_extended/contextlib_suppress_multiple
# origin: languages/python/tests/python/test_context_managers_extended.rs
# vybe-test-mode: compile

from contextlib import suppress
with suppress(ValueError, TypeError):
 pass
