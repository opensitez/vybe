# vybe-test: python/for_else_core/for_else_predicate_fails_triggers_break
# origin: languages/python/tests/python/test_for_else_core.rs

nums = [2, -1, 8]
for n in nums:
 if n < 0:
  print('bad')
  break
else:
 print('ok')
