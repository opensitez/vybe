# vybe-test: python/while_loops/while_with_boolean_flag
# origin: languages/python/tests/python/test_while_loops.rs

running = True
c = 0
while running:
 c += 1
 if c == 3:
  running = False
print(c)
