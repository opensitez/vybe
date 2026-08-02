# vybe-test: python/for_while_extended/for_else_break_skips
# origin: languages/python/tests/python/test_for_while_extended.rs

for i in range(3):
 if i == 1:
  break
else:
 print('else')
print('done')
