# vybe-test: python/match_extended/match_break_in_loop
# origin: languages/python/tests/python/test_match_extended.rs

for i in range(3):
 match i:
  case 1:
   print('hit')
   break
