# vybe-test: python/pythonic_idioms/enumerate_in_for_loop
# origin: languages/python/tests/python/test_pythonic_idioms.rs

for i, v in enumerate(['a', 'b']):
 if i == 1:
  print(v)
