# vybe-test: python/while_loops/while_else_runs_when_not_broken
# origin: languages/python/tests/python/test_while_loops.rs

n = 2
while n:
 n -= 1
else:
 print('done')
