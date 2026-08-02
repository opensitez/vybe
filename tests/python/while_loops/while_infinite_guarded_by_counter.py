# vybe-test: python/while_loops/while_infinite_guarded_by_counter
# origin: languages/python/tests/python/test_while_loops.rs

c = 0
while True:
 c += 1
 if c == 4:
  break
print(c)
