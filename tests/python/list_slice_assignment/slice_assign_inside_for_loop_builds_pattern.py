# vybe-test: python/list_slice_assignment/slice_assign_inside_for_loop_builds_pattern
# origin: languages/python/tests/python/test_list_slice_assignment.rs

xs = [0, 0, 0, 0]
for i in range(2):
 xs[i:i+1] = [i + 1]
print(xs)
