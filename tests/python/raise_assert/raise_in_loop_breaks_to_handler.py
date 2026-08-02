# vybe-test: python/raise_assert/raise_in_loop_breaks_to_handler
# origin: languages/python/tests/python/test_raise_assert.rs

for i in range(2):
 try:
  if i:
   raise ValueError('loop')
 except ValueError:
  print('caught')
