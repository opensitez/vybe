# vybe-test: python/for_else_core/for_else_sorted_unique_check
# origin: languages/python/tests/python/test_for_else_core.rs

xs = [1, 2, 3]
for i in range(1, len(xs)):
 if xs[i] < xs[i-1]:
  print('unsorted')
  break
else:
 print('sorted')
