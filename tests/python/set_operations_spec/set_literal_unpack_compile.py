# vybe-test: python/set_operations_spec/set_literal_unpack_compile
# origin: languages/python/tests/python/test_set_operations_spec.rs
# vybe-test-mode: compile

a = {2, 3}
b = {1, *a, 4}
