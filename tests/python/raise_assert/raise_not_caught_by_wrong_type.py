# vybe-test: python/raise_assert/raise_not_caught_by_wrong_type
# origin: languages/python/tests/python/test_raise_assert.rs

try:
 try:
  raise TypeError('t')
 except ValueError:
  print('wrong')
except TypeError:
 print('right')
