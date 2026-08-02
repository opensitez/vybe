# vybe-test: python/match_extended/match_sequence_or
# origin: languages/python/tests/python/test_match_extended.rs
# vybe-test-mode: compile

match x:
 case [1, 2] | [3, 4]:
  pass
