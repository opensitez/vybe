# vybe-test: python/for_else_core/for_else_break_on_match_skips_else
# origin: languages/python/tests/python/test_for_else_core.rs

for x in [1, 2, 3]:
 if x == 2:
  print('hit')
  break
else:
 print('else')
