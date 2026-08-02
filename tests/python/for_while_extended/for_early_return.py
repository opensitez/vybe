# vybe-test: python/for_while_extended/for_early_return
# origin: languages/python/tests/python/test_for_while_extended.rs

def f():
 for i in range(5):
  if i == 2:
   return i
 return -1
print(f())
