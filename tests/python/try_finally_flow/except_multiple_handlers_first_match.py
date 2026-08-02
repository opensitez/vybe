# vybe-test: python/try_finally_flow/except_multiple_handlers_first_match
# origin: languages/python/tests/python/test_try_finally_flow.rs

try:
 raise TypeError('t')
except ValueError:
 print('v')
except TypeError:
 print('t')
