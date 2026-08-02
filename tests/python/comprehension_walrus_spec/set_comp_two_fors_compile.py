# vybe-test: python/comprehension_walrus_spec/set_comp_two_fors_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

s = {(i, j) for i in range(2) for j in range(2)}
