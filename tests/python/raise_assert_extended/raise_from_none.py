# vybe-test: python/raise_assert_extended/raise_from_none
# origin: languages/python/tests/python/test_raise_assert_extended.rs
# vybe-test-mode: compile

try:
 raise ValueError() from None
except ValueError:
 pass
