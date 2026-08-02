# vybe-test: python/nested_loop_control/break_outer_via_flag_pattern
# origin: languages/python/tests/python/test_nested_loop_control.rs

out = []
for i in range(3):
 for j in range(3):
  if i == 1 and j == 1:
   out.append('stop')
   break
  out.append(i)
 else:
  continue
 break
print(out)
