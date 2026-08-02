# vybe-test: python/generator_extended/generator_enumerate_yield
# origin: languages/python/tests/python/test_generator_extended.rs

def g():
 for i, v in enumerate(['a', 'b']):
  yield i, v
print(list(g()))
