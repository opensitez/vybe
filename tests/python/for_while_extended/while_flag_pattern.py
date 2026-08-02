# vybe-test: python/for_while_extended/while_flag_pattern
# origin: languages/python/tests/python/test_for_while_extended.rs

done = False
n = 0
while not done:
 n += 1
 if n >= 2:
  done = True
print(n)
