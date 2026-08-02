# vybe-test: python/match_extended/match_guard_false_fallthrough
# origin: languages/python/tests/python/test_match_extended.rs

x = 2
match x:
 case n if n > 3:
  print('big')
 case _:
  print('other')
