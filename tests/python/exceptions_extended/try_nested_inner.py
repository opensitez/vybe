# vybe-test: python/exceptions_extended/try_nested_inner
# origin: languages/python/tests/python/test_exceptions_extended.rs

try:
 try:
  1/0
 except ZeroDivisionError:
  print('inner')
except:
 print('outer')
