# vybe-test: python/exceptions_extended/except_exception_chaining_context
# origin: languages/python/tests/python/test_exceptions_extended.rs
# vybe-test-mode: compile

try:
 1/0
except ZeroDivisionError as e:
 raise ValueError() from e
