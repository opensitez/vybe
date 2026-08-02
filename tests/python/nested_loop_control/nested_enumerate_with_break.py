# vybe-test: python/nested_loop_control/nested_enumerate_with_break
# origin: languages/python/tests/python/test_nested_loop_control.rs

out = []
for i, row in enumerate([[1, 2], [3, 4]]):
 for j, v in enumerate(row):
  if v == 3:
   out.append(i * 10 + j)
   break
print(out)
