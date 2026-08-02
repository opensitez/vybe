# vybe-test: python/for_else_core/for_else_finds_no_match_sets_flag
# origin: languages/python/tests/python/test_for_else_core.rs

found = False
for x in [1, 3, 5]:
 if x == 2:
  found = True
  break
else:
 print('missing')
print(found)
