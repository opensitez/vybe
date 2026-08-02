# vybe-test: python/slicing_extended/slice_del_negative_step
# origin: languages/python/tests/python/test_slicing_extended.rs
# vybe-test-mode: compile

a = [1,2,3,4]
del a[::-2]
