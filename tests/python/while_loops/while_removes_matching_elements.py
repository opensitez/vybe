# vybe-test: python/while_loops/while_removes_matching_elements
# origin: languages/python/tests/python/test_while_loops.rs

xs = [1, 2, 3, 2, 1]
i = 0
while i < len(xs):
 if xs[i] == 2:
  xs.pop(i)
 else:
  i += 1
print(xs.count(2))
