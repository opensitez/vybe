# vybe-test: python/for_else_core/for_else_inner_break_does_not_trigger_outer_else
# origin: languages/python/tests/python/test_for_else_core.rs

for i in range(2):
 for j in range(3):
  if j == 1:
   break
 else:
  print('inner')
else:
 print('outer')
