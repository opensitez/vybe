# vybe-test: python/try_except_core/try_except_nested_inner_caught
# origin: languages/python/tests/python/test_try_except_core.rs

try:
 try:
  1/0
 except ZeroDivisionError:
  print('inner')
except:
 print('outer')
