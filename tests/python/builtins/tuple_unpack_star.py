# vybe-test: python/builtins/tuple_unpack_star
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

a = (1, 2)
b = (*a, 3, 4)
