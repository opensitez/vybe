# vybe-test: python/try_finally_flow/try_nested_inner_except
# origin: languages/python/tests/python/test_try_finally_flow.rs

try:
 try:
  1/0
 except ZeroDivisionError:
  print('inner')
except:
 print('outer')
