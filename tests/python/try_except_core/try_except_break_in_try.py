# vybe-test: python/try_except_core/try_except_break_in_try
# origin: languages/python/tests/python/test_try_except_core.rs

for i in range(3):
 try:
  if i == 1:
   break
 except:
  pass
print(i)
