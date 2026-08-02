# vybe-test: python/for_else_core/for_else_validates_all_predicates
# origin: languages/python/tests/python/test_for_else_core.rs

nums = [2, 4, 8]
for n in nums:
 if n < 0:
  break
else:
 print('ok')
