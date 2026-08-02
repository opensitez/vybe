# vybe-test: python/match_core/match_continue_in_loop
# origin: languages/python/tests/python/test_match_core.rs

out = []
for v in [1, 2, 3]:
 match v:
  case 2:
   continue
 out.append(v)
print(out)
