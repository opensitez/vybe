# vybe-test: python/for_while_extended/while_continue
# origin: languages/python/tests/python/test_for_while_extended.rs

i = 0
while i < 4:
 i += 1
 if i % 2 == 0:
  continue
 print(i)
 break
