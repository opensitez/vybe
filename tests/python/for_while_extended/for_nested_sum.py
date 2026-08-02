# vybe-test: python/for_while_extended/for_nested_sum
# origin: languages/python/tests/python/test_for_while_extended.rs

s = 0
for i in range(2):
 for j in range(2):
  s += i + j
print(s)
