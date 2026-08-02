# vybe-test: python/for_while_extended/while_walrus
# origin: languages/python/tests/python/test_for_while_extended.rs

data = [1, 2, 3]
i = 0
while (v := data[i] if i < len(data) else None) is not None:
 print(v)
 i += 1
 if i > 0:
  break
