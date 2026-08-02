# vybe-test: python/set_operations_spec/set_isdisjoint_compile
# origin: languages/python/tests/python/test_set_operations_spec.rs
# vybe-test-mode: compile

a = {1, 2}
b = {3, 4}
ok = a.isdisjoint(b)
