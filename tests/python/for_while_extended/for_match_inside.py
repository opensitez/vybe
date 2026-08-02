# vybe-test: python/for_while_extended/for_match_inside
# origin: languages/python/tests/python/test_for_while_extended.rs

for x in [1, 2]:
 match x:
  case 1:
   print('one')
   break
