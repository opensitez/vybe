# vybe-test: python/match_extended/match_continue_in_loop
# origin: languages/python/tests/python/test_match_extended.rs

for i in range(3):
 match i:
  case 0:
   continue
 print(i)
