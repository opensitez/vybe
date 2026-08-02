# vybe-test: python/nested_loop_control/nested_loop_counts_pairs_skip_diagonal
# origin: languages/python/tests/python/test_nested_loop_control.rs

n = 0
for i in range(3):
 for j in range(3):
  if i == j:
   continue
  n += 1
print(n)
