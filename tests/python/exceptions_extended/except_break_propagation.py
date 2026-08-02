# vybe-test: python/exceptions_extended/except_break_propagation
# origin: languages/python/tests/python/test_exceptions_extended.rs

for i in range(1):
 try:
  break
 except:
  pass
print('done')
