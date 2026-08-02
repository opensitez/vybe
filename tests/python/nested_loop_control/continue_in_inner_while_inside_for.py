# vybe-test: python/nested_loop_control/continue_in_inner_while_inside_for
# origin: languages/python/tests/python/test_nested_loop_control.rs

out = []
for x in [1, 2]:
 y = 0
 while y < 4:
  y += 1
  if y == 2:
   continue
  out.append(x + y)
print(out)
