# vybe-test: python/for_else_core/for_else_first_duplicate_finder
# origin: languages/python/tests/python/test_for_else_core.rs

xs = [1, 2, 3, 2, 4]
seen = set()
for x in xs:
 if x in seen:
  print('dup', x)
  break
 seen.add(x)
else:
 print('unique')
