# vybe-test: python/generators_core/generator_fibonacci_style
# origin: languages/python/tests/python/test_generators_core.rs

def fib():
 a, b = 0, 1
 while a < 10:
  yield a
  a, b = b, a + b
print(list(fib()))
