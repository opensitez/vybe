# vybe-test: python/match_extended/match_in_loop
# origin: languages/python/tests/python/test_match_extended.rs

out = []
for v in [1, 2, 3]:
 match v:
  case 2:
   out.append('two')
print(out)
