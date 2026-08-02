# vybe-test: python/enumerate_start_index/enumerate_nested_loop_outer_index
# origin: languages/python/tests/python/test_enumerate_start_index.rs

out = []
for i, row in enumerate([[1], [2, 3]]):
 out.append(i)
print(out)
