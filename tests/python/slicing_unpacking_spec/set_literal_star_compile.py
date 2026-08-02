# vybe-test: python/slicing_unpacking_spec/set_literal_star_compile
# origin: languages/python/tests/python/test_slicing_unpacking_spec.rs
# vybe-test-mode: compile

a = {2, 3}
b = {1, *a, 4}
