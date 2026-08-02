# vybe-test: python/generators_core/generator_count_with_sentinel
# origin: languages/python/tests/python/test_generators_core.rs

def count(n):
 while n > 0:
  yield n
  n -= 1
print(list(count(3)))
