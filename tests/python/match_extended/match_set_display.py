# vybe-test: python/match_extended/match_set_display
# origin: languages/python/tests/python/test_match_extended.rs

x = {1, 2}
match x:
 case {1, 2}:
  print('set')
 case _:
  print('no')
