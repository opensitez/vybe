# vybe-test: python/nested_loop_control/while_true_inner_break_counter
# origin: languages/python/tests/python/test_nested_loop_control.rs

n = 0
while True:
 while True:
  n += 1
  if n == 3:
   break
 break
print(n)
