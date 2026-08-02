# vybe-test: python/try_except_core/try_except_reraise_not_caught_by_wrong_type
# origin: languages/python/tests/python/test_try_except_core.rs

try:
 try:
  raise ValueError('x')
 except TypeError:
  print('wrong')
except ValueError:
 print('right')
