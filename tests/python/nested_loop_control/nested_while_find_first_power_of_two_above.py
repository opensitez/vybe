# vybe-test: python/nested_loop_control/nested_while_find_first_power_of_two_above
# origin: languages/python/tests/python/test_nested_loop_control.rs

v = 1
while v < 20:
 p = 1
 while p < v:
  p *= 2
 if p == v:
  print(v)
  break
 v += 1
else:
 print('none')
