# vybe-test: python/match_core/match_first_case_wins
# origin: languages/python/tests/python/test_match_core.rs

x = 1
match x:
 case 1:
  print('first')
 case 1:
  print('second')
