# vybe-test: python/try_except_core/try_except_multiple_handlers_second_match
# origin: languages/python/tests/python/test_try_except_core.rs

try:
 raise ValueError('v')
except TypeError:
 print('t')
except ValueError:
 print('v')
