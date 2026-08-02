# vybe-test: python/while_loops/while_spins_on_flag_then_clears
# origin: languages/python/tests/python/test_while_loops.rs

flag = True
c = 0
while flag:
 c += 1
 if c == 1:
  flag = False
print(c)
