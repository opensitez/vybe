# vybe-test: python/walrus_extended/walrus_for_else
# origin: languages/python/tests/python/test_walrus_extended.rs

for i in range(1):
 if (x := i) == 0:
  print('ok')
else:
 print('skip')
