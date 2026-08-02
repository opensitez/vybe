# vybe-test: python/nested_loop_control/nested_continue_only_affects_inner_index
# origin: languages/python/tests/python/test_nested_loop_control.rs

pairs = []
for a in range(2):
 for b in range(4):
  if b % 2 == 1:
   continue
  pairs.append((a, b))
print(pairs)
