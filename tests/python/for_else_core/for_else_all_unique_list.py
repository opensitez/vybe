# vybe-test: python/for_else_core/for_else_all_unique_list
# origin: languages/python/tests/python/test_for_else_core.rs

xs = [1, 2, 3]
seen = set()
for x in xs:
 if x in seen:
  print('dup')
  break
 seen.add(x)
else:
 print('unique')
