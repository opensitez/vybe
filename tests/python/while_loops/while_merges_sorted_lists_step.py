# vybe-test: python/while_loops/while_merges_sorted_lists_step
# origin: languages/python/tests/python/test_while_loops.rs

a = [1, 3]
b = [2, 4]
i = j = 0
out = []
while i < len(a) and j < len(b):
 if a[i] < b[j]:
  out.append(a[i])
  i += 1
 else:
  out.append(b[j])
  j += 1
print(out[2])
