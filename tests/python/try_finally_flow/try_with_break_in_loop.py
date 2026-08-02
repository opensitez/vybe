# vybe-test: python/try_finally_flow/try_with_break_in_loop
# origin: languages/python/tests/python/test_try_finally_flow.rs

for _ in range(3):
 try:
  print('x')
  break
 finally:
  print('y')
