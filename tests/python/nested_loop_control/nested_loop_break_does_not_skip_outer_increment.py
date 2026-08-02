# vybe-test: python/nested_loop_control/nested_loop_break_does_not_skip_outer_increment
# origin: languages/python/tests/python/test_nested_loop_control.rs

total = 0
for i in range(3):
 for j in range(5):
  if j == 2:
   break
  total += 1
print(total)
