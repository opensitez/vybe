# vybe-test: python/while_loops/while_nested_multiplies_counters
# origin: languages/python/tests/python/test_while_loops.rs

i = 0
prod = 1
while i < 3:
 j = 0
 while j < 2:
  prod *= 2
  j += 1
 i += 1
print(prod)
