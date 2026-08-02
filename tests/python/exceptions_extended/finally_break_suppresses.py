# vybe-test: python/exceptions_extended/finally_break_suppresses
# origin: languages/python/tests/python/test_exceptions_extended.rs

for i in range(1):
 try:
  break
 finally:
  pass
print('after')
