# vybe-test: python/slicing_extended/slice_assign_negative_step
# origin: languages/python/tests/python/test_slicing_extended.rs
# vybe-test-mode: compile

a = [0,1,2,3,4]
a[4:0:-2] = [9,9]
