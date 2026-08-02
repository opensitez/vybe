# vybe-test: python/for_while_extended/for_continue_skip
# origin: languages/python/tests/python/test_for_while_extended.rs

for i in range(4):
 if i % 2 == 0:
  continue
 print(i)
 break
