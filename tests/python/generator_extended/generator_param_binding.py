# vybe-test: python/generator_extended/generator_param_binding
# origin: languages/python/tests/python/test_generator_extended.rs

def g(n):
 for i in range(n):
  yield i
print(list(g(3)))
