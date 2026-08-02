# vybe-test: python/for_else_core/for_else_unsorted_detected
# origin: languages/python/tests/python/test_for_else_core.rs

xs = [1, 3, 2]
for i in range(1, len(xs)):
 if xs[i] < xs[i-1]:
  print('unsorted')
  break
else:
 print('sorted')
