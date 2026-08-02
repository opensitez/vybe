# vybe-test: python/while_loops/while_accumulates_factorial
# origin: languages/python/tests/python/test_while_loops.rs

n = 5
f = 1
while n > 1:
 f *= n
 n -= 1
print(f)
