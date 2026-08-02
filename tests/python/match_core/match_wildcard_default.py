# vybe-test: python/match_core/match_wildcard_default
# origin: languages/python/tests/python/test_match_core.rs

x = 9
match x:
 case 1:
  print('no')
 case _:
  print('other')
