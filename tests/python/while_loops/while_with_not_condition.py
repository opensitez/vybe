# vybe-test: python/while_loops/while_with_not_condition
# origin: languages/python/tests/python/test_while_loops.rs

done = False
c = 0
while not done:
 c += 1
 done = c == 2
print(c)
