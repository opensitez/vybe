# vybe-test: python/set_operations_spec/set_issubset_compile
# origin: languages/python/tests/python/test_set_operations_spec.rs
# vybe-test-mode: compile

a = {1, 2}
b = {1, 2, 3}
ok = a.issubset(b)
