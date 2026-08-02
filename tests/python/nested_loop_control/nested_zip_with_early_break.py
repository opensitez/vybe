# vybe-test: python/nested_loop_control/nested_zip_with_early_break
# origin: languages/python/tests/python/test_nested_loop_control.rs

out = []
for a, b in zip([1, 2, 3], ['x', 'y', 'z']):
 for c in range(2):
  if c == 1:
   break
  out.append(str(a) + b + str(c))
print(out)
