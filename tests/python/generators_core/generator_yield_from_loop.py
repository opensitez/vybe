# vybe-test: python/generators_core/generator_yield_from_loop
# origin: languages/python/tests/python/test_generators_core.rs

def g():
 for i in range(3):
  yield i
print(list(g()))
