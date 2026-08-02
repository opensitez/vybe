# vybe-test: python/for_else_core/for_else_with_return_in_function
# origin: languages/python/tests/python/test_for_else_core.rs

def f():
 for x in range(2):
  if x == 5:
   break
 else:
  return 'done'
 return 'skip'
print(f())
