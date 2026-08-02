# vybe-test: python/builtins/list_unpack_multiple
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

x = [1, 2]
y = [3, 4]
z = [*x, *y]
