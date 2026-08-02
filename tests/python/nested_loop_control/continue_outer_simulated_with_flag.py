# vybe-test: python/nested_loop_control/continue_outer_simulated_with_flag
# origin: languages/python/tests/python/test_nested_loop_control.rs

out = []
for i in range(4):
 skip = False
 for j in range(2):
  if i == 2:
   skip = True
   break
 if skip:
  continue
 out.append(i)
print(out)
