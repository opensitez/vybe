# vybe-test: python/while_loops/do_while_emulated_with_first_run
# origin: languages/python/tests/python/test_while_loops.rs

n = 0
while True:
 n += 1
 print(n)
 if n >= 2:
  break
