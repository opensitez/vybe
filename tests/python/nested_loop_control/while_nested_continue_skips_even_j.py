# vybe-test: python/nested_loop_control/while_nested_continue_skips_even_j
# origin: languages/python/tests/python/test_nested_loop_control.rs

i = 0
out = []
while i < 2:
 j = 0
 while j < 4:
  j += 1
  if j % 2 == 0:
   continue
  out.append(i * 10 + j)
 i += 1
print(out)
