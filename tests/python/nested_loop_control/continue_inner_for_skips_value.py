# vybe-test: python/nested_loop_control/continue_inner_for_skips_value
# origin: languages/python/tests/python/test_nested_loop_control.rs

out = []
for i in range(2):
 for j in range(3):
  if j == 1:
   continue
  out.append(j)
print(out)
