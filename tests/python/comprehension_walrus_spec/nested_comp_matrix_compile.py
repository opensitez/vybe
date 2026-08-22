# vybe-test: python/comprehension_walrus_spec/nested_comp_matrix_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs

matrix = [[i * j for j in range(3)] for i in range(3)]
