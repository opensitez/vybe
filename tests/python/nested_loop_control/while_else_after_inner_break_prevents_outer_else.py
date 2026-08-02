# vybe-test: python/nested_loop_control/while_else_after_inner_break_prevents_outer_else
# origin: languages/python/tests/python/test_nested_loop_control.rs

n = 0
while n < 2:
 m = 0
 while m < 2:
  if m == 1:
   break
  m += 1
 else:
  print('inner')
  n += 1
  continue
 break
else:
 print('outer')
