# vybe-test: python/for_else_core/for_else_prime_check_small_prime
# origin: languages/python/tests/python/test_for_else_core.rs

n = 7
for d in range(2, n):
 if n % d == 0:
  print('composite')
  break
else:
 print('prime')
