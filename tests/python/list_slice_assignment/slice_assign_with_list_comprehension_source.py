# vybe-test: python/list_slice_assignment/slice_assign_with_list_comprehension_source
# origin: languages/python/tests/python/test_list_slice_assignment.rs

xs = [0, 0, 0, 0]
xs[1:3] = [n * 10 for n in range(2)]
print(xs)
