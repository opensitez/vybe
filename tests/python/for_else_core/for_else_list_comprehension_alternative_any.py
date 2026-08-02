# vybe-test: python/for_else_core/for_else_list_comprehension_alternative_any
# origin: languages/python/tests/python/test_for_else_core.rs

xs = [2, 4, 6]
for x in xs:
 if x % 2 == 1:
  print('odd')
  break
else:
 print('all even')
