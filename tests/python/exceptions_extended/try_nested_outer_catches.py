# vybe-test: python/exceptions_extended/try_nested_outer_catches
# origin: languages/python/tests/python/test_exceptions_extended.rs

try:
 try:
  raise KeyError()
 except TypeError:
  print('no')
except KeyError:
 print('outer')
