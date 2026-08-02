# vybe-test: python/while_loops/while_else_skipped_on_break
# origin: languages/python/tests/python/test_while_loops.rs

n = 5
while n:
 n -= 1
 if n == 2:
  break
else:
 print('no')
print('yes')
