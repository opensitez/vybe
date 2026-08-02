# vybe-test: python/nested_loop_control/for_else_skipped_when_inner_breaks
# origin: languages/python/tests/python/test_nested_loop_control.rs

for i in range(2):
 for j in range(2):
  if j == 1:
   break
else:
 print('no')
