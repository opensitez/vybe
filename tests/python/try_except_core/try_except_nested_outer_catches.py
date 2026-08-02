# vybe-test: python/try_except_core/try_except_nested_outer_catches
# origin: languages/python/tests/python/test_try_except_core.rs

try:
 try:
  raise KeyError
 except ValueError:
  print('inner')
except KeyError:
 print('outer')
