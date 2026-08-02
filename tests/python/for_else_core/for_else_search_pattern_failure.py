# vybe-test: python/for_else_core/for_else_search_pattern_failure
# origin: languages/python/tests/python/test_for_else_core.rs

target = 4
for n in [1, 3, 5]:
 if n == target:
  print('found')
  break
else:
 print('not found')
