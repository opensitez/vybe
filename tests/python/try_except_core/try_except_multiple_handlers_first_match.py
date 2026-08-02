# vybe-test: python/try_except_core/try_except_multiple_handlers_first_match
# origin: languages/python/tests/python/test_try_except_core.rs

try:
 raise TypeError('t')
except ValueError:
 print('v')
except TypeError:
 print('t')
