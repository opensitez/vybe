# vybe-test: python/generator_extended/generator_infinite_take
# origin: languages/python/tests/python/test_generator_extended.rs

def count():
 n = 0
 while True:
  yield n
  n += 1
print([next(count()) for _ in range(3)])
