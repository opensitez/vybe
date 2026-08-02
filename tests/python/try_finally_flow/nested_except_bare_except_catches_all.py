# vybe-test: python/try_finally_flow/nested_except_bare_except_catches_all
# origin: languages/python/tests/python/test_try_finally_flow.rs

try:
 try:
  raise ValueError
 except:
  print('inner')
except:
 print('outer')
