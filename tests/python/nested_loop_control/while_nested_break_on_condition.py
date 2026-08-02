# vybe-test: python/nested_loop_control/while_nested_break_on_condition
# origin: languages/python/tests/python/test_nested_loop_control.rs

i = 0
out = []
while i < 3:
 j = 0
 while j < 3:
  if j == 2:
   break
  out.append(i * 10 + j)
  j += 1
 i += 1
print(out)
