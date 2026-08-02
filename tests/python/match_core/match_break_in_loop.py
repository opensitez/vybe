# vybe-test: python/match_core/match_break_in_loop
# origin: languages/python/tests/python/test_match_core.rs

for v in [1, 2]:
 match v:
  case 2:
   print('stop')
   break
