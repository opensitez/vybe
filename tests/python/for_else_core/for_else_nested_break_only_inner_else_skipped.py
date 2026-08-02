# vybe-test: python/for_else_core/for_else_nested_break_only_inner_else_skipped
# origin: languages/python/tests/python/test_for_else_core.rs

for i in range(2):
 for j in range(2):
  if j == 0:
   break
 else:
  print('inner-else')
else:
 print('outer-else')
