# vybe-test: python/exceptions_extended/except_bare_except
# origin: languages/python/tests/python/test_exceptions_extended.rs
# vybe-test-mode: compile

try:
 raise ValueError()
except:
 pass
