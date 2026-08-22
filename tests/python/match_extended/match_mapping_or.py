# vybe-test: python/match_extended/match_mapping_or
# origin: languages/python/tests/python/test_match_extended.rs
d = 1

match d:
 case {'a': 1} | {'b': 2}:
  pass
