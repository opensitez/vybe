# vybe-test: python/for_while_extended/for_nested_else
# origin: languages/python/tests/python/test_for_while_extended.rs

for i in range(1):
 for j in range(1):
  pass
 else:
  print('inner')
