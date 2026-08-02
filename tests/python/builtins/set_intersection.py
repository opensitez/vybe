# vybe-test: python/builtins/set_intersection
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

a = {1, 2, 3}
b = {2, 3, 4}
c = a.intersection(b)
