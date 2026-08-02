# vybe-test: python/builtins/slice_assign_basic
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

a = [1, 2, 3, 4, 5]
a[1:3] = [10, 20]
