# vybe-test: python/for_else_core/for_else_membership_all_present
# origin: languages/python/tests/python/test_for_else_core.rs

need = 'ae'
word = 'cat'
for ch in need:
 if ch not in word:
  print('missing')
  break
else:
 print('has all')
