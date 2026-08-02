# vybe-test: python/nested_loop_control/nested_for_else_break_on_last_outer_iteration
# origin: languages/python/tests/python/test_nested_loop_control.rs

out = []
for i in range(3):
 for j in range(2):
  if i == 2 and j == 1:
   break
  out.append(i)
 else:
  out.append(99)
print(out)
