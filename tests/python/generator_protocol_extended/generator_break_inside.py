# vybe-test: python/generator_protocol_extended/generator_break_inside
# origin: languages/python/tests/python/test_generator_protocol_extended.rs

def g():
 for i in range(10):
  if i == 3:
   break
  yield i
print(list(g()))
