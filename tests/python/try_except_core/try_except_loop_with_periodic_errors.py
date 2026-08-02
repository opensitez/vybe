# vybe-test: python/try_except_core/try_except_loop_with_periodic_errors
# origin: languages/python/tests/python/test_try_except_core.rs

count = 0
for i in range(3):
 try:
  if i == 1:
   raise ValueError
  count += 1
 except ValueError:
  pass
print(count)
