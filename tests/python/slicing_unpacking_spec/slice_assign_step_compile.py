# vybe-test: python/slicing_unpacking_spec/slice_assign_step_compile
# origin: languages/python/tests/python/test_slicing_unpacking_spec.rs
# vybe-test-mode: compile

x = [0, 1, 2, 3, 4, 5]
x[::2] = [9, 9, 9]
