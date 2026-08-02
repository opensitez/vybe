# vybe-test: python/while_loops/while_else_not_run_on_return_emulated
# origin: languages/python/tests/python/test_while_loops.rs

n = 1
while n:
 print('loop')
 break
else:
 print('else')
print('after')
