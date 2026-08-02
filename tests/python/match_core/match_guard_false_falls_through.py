# vybe-test: python/match_core/match_guard_false_falls_through
# origin: languages/python/tests/python/test_match_core.rs

x = 1
match x:
 case n if n > 5:
  print('big')
 case _:
  print('small')
