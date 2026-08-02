# vybe-test: python/for_while_extended/while_else_break_skips
# origin: languages/python/tests/python/test_for_while_extended.rs

i = 0
while i < 3:
 if i == 1:
  break
 i += 1
else:
 print('else')
print('fin')
