# vybe-test: python/match_extended/match_capture_walrus
# origin: languages/python/tests/python/test_match_extended.rs
# vybe-test-mode: compile

match x:
 case n if (d := n // 2):
  pass
