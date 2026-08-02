# vybe-test: python/match_core/match_no_case_raises
# origin: languages/python/tests/python/test_match_core.rs

x = 1
match x:
 case 2:
  print('no')
try:
 match x:
  case 2:
   pass
except:
 print('unmatched')
