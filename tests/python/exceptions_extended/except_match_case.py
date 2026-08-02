# vybe-test: python/exceptions_extended/except_match_case
# origin: languages/python/tests/python/test_exceptions_extended.rs
# vybe-test-mode: compile

try:
 raise ValueError()
except ValueError:
 match 1:
  case 1:
   pass
