# vybe-test: python/while_loops/while_break_exits_immediately
# origin: languages/python/tests/python/test_while_loops.rs

n = 10
while n:
 n -= 1
 if n == 7:
  break
print(n)
