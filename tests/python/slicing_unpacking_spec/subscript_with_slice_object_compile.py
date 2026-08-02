# vybe-test: python/slicing_unpacking_spec/subscript_with_slice_object_compile
# origin: languages/python/tests/python/test_slicing_unpacking_spec.rs
# vybe-test-mode: compile

x = [1, 2, 3, 4, 5]
sl = slice(1, 4)
y = x[sl]
