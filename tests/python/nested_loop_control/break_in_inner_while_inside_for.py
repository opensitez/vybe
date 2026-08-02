# vybe-test: python/nested_loop_control/break_in_inner_while_inside_for
# origin: languages/python/tests/python/test_nested_loop_control.rs

out = []
for x in [1, 2]:
 y = 0
 while y < 5:
  if y == 2:
   break
  out.append(x * 10 + y)
  y += 1
print(out)
