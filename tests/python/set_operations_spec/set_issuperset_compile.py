# vybe-test: python/set_operations_spec/set_issuperset_compile
# origin: languages/python/tests/python/test_set_operations_spec.rs
# vybe-test-mode: compile

a = {1, 2, 3}
b = {1, 2}
ok = a.issuperset(b)
