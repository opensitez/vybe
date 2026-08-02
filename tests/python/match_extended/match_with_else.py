# vybe-test: python/match_extended/match_with_else
# origin: languages/python/tests/python/test_match_extended.rs

x = 5
match x:
 case 1:
  print('one')
 case _:
  print('done')
