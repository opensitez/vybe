# vybe-test: python/generator_extended/generator_yield_in_loop
# origin: languages/python/tests/python/test_generator_extended.rs

def g():
 for i in range(3):
  yield i
print(list(g()))
