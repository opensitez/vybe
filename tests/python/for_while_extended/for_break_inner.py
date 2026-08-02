# vybe-test: python/for_while_extended/for_break_inner
# origin: languages/python/tests/python/test_for_while_extended.rs

for i in range(3):
 for j in range(3):
  if j == 1:
   break
 print(i, j)
 break
