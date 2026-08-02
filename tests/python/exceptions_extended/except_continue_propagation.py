# vybe-test: python/exceptions_extended/except_continue_propagation
# origin: languages/python/tests/python/test_exceptions_extended.rs

for i in range(2):
 try:
  if i == 0:
   continue
  print(i)
 except:
  pass
