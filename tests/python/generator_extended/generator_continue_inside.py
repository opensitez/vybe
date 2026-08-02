# vybe-test: python/generator_extended/generator_continue_inside
# origin: languages/python/tests/python/test_generator_extended.rs

def g():
 for i in range(4):
  if i % 2 == 0:
   continue
  yield i
print(list(g()))
