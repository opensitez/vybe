# vybe-test: python/walrus_extended/walrus_list_append_in_loop
# origin: languages/python/tests/python/test_walrus_extended.rs

out = []
for i in range(3):
 if (v := i * i) >= 0:
  out.append(v)
print(out)
