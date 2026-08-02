# vybe-test: python/match_extended/match_first_branch_wins
# origin: languages/python/tests/python/test_match_extended.rs

x = 1
match x:
 case 1:
  print('first')
 case 1:
  print('second')
