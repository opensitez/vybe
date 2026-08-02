# vybe-test: python/for_else_core/for_else_with_continue_still_runs
# origin: languages/python/tests/python/test_for_else_core.rs

for x in range(3):
 if x == 1:
  continue
else:
 print('done')
