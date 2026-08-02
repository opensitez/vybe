# vybe-test: python/zip_patterns/zip_list_comprehension_sum_products
# origin: languages/python/tests/python/test_zip_patterns.rs

sum(a * b for a, b in zip([1, 2, 3], [4, 5, 6]))
