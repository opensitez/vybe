# vybe-test: python/match_extended/match_sequence_length_mismatch
# origin: languages/python/tests/python/test_match_extended.rs

x = [1, 2]
match x:
 case [a, b, c]:
  print('no')
 case _:
  print('fallback')
