# vybe-test: python/builtins/set_union
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

a = {1, 2}
b = {3, 4}
c = a.union(b)
