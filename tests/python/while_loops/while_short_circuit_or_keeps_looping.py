# vybe-test: python/while_loops/while_short_circuit_or_keeps_looping
# origin: languages/python/tests/python/test_while_loops.rs

a = 0
b = 1
c = 0
while a or b:
 c += 1
 b = 0
 if c == 2:
  break
print(c)
