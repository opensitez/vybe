# vybe-test: python/raise_assert_extended/raise_chained_context
# origin: languages/python/tests/python/test_raise_assert_extended.rs
# vybe-test-mode: compile

try:
 1/0
except ZeroDivisionError as e:
 raise ValueError() from e
