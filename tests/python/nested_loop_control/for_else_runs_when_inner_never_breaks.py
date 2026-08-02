# vybe-test: python/nested_loop_control/for_else_runs_when_inner_never_breaks
# origin: languages/python/tests/python/test_nested_loop_control.rs

for i in range(2):
 for j in range(2):
  pass
else:
 print('ok')
