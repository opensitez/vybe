# vybe-test: python/while_loops/while_modulo_cycle_detects_period
# origin: languages/python/tests/python/test_while_loops.rs

n = 1
steps = 0
while steps < 4:
 n = (n * 3) % 7
 steps += 1
print(n)
