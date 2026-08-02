# vybe-test: python/nested_loop_control/triple_nested_accumulate_until_limit
# origin: languages/python/tests/python/test_nested_loop_control.rs

n = 0
for a in range(2):
 for b in range(2):
  for c in range(2):
   n += 1
print(n)
