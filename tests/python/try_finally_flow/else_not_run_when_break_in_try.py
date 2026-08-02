# vybe-test: python/try_finally_flow/else_not_run_when_break_in_try
# origin: languages/python/tests/python/test_try_finally_flow.rs

for _ in range(1):
 try:
  break
 else:
  print('else')
print('done')
