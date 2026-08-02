# vybe-test: python/for_while_extended/for_else_return_in_body
# origin: languages/python/tests/python/test_for_while_extended.rs

def f():
 for i in range(1):
  return i
 else:
  return -1
print(f())
