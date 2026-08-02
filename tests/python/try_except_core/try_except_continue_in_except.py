# vybe-test: python/try_except_core/try_except_continue_in_except
# origin: languages/python/tests/python/test_try_except_core.rs

out = []
for i in range(3):
 try:
  if i == 1:
   raise ValueError
  out.append(i)
 except ValueError:
  continue
print(out)
