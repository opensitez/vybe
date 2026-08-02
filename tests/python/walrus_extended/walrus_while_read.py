# vybe-test: python/walrus_extended/walrus_while_read
# origin: languages/python/tests/python/test_walrus_extended.rs

s = 'abc'
i = 0
while (c := s[i] if i < len(s) else ''):
 i += 1
 if i >= 2:
  break
print(c)
