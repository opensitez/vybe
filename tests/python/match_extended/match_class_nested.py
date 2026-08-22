# vybe-test: python/match_extended/match_class_nested
# origin: languages/python/tests/python/test_match_extended.rs
o = 1

class A:
 pass
match o:
 case A():
  pass
