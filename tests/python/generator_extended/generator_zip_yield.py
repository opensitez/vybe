# vybe-test: python/generator_extended/generator_zip_yield
# origin: languages/python/tests/python/test_generator_extended.rs

def g():
 for a, b in zip([1, 2], [3, 4]):
  yield a + b
print(list(g()))
