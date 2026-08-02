# vybe-test: python/generator_protocol_extended/generator_param
# origin: languages/python/tests/python/test_generator_protocol_extended.rs

def g(n):
 for i in range(n):
  yield i
print(list(g(3)))
