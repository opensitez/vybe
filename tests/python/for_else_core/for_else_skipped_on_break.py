# vybe-test: python/for_else_core/for_else_skipped_on_break
# origin: languages/python/tests/python/test_for_else_core.rs

for x in range(5):
 if x == 2:
  break
else:
 print('no')
print('yes')
