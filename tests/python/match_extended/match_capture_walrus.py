# vybe-test: python/match_extended/match_capture_walrus
# origin: languages/python/tests/python/test_match_extended.rs
x = 1

match x:
 case n if (d := n // 2):
  pass
