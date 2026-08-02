# vybe-test: python/for_while_extended/while_nested_break
# origin: languages/python/tests/python/test_for_while_extended.rs

for i in range(2):
 j = 0
 while j < 3:
  if j == 1:
   break
  j += 1
 print(i, j)
 break
