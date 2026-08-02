# vybe-test: python/comprehension_walrus_spec/nested_comp_matrix_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

matrix = [[i * j for j in range(3)] for i in range(3)]
