# vybe-test: python/set_comprehension/set_comp_zip_pairs_as_tuples
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({(a, b) for a, b in zip([1, 2], [3, 4])})
