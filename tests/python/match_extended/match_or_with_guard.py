# vybe-test: python/match_extended/match_or_with_guard
# origin: languages/python/tests/python/test_match_extended.rs

x = 3
match x:
 case 1 | 2:
  print('low')
 case 3 | 4 if x == 3:
  print('mid')
