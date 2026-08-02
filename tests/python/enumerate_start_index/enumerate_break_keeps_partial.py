# vybe-test: python/enumerate_start_index/enumerate_break_keeps_partial
# origin: languages/python/tests/python/test_enumerate_start_index.rs

out = []
for i, v in enumerate([1, 2, 3]):
 if i == 1:
  break
 out.append(i)
print(out)
