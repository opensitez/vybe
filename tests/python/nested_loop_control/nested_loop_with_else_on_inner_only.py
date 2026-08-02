# vybe-test: python/nested_loop_control/nested_loop_with_else_on_inner_only
# origin: languages/python/tests/python/test_nested_loop_control.rs

out = []
for i in range(2):
 for j in range(2):
  out.append(i + j)
 else:
  out.append(9)
print(out)
