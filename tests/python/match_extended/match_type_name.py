# vybe-test: python/match_extended/match_type_name
# origin: languages/python/tests/python/test_match_extended.rs

x = 1
match x:
 case int():
  print('int')
 case _:
  print('other')
