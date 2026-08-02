# vybe-test: python/for_else_core/for_else_while_emulated_via_for_range
# origin: languages/python/tests/python/test_for_else_core.rs

n = 3
for _ in range(100):
 n -= 1
 if n <= 0:
  break
else:
 print('limit')
print(n)
