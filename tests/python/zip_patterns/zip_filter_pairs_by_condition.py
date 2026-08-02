# vybe-test: python/zip_patterns/zip_filter_pairs_by_condition
# origin: languages/python/tests/python/test_zip_patterns.rs

[(a, b) for a, b in zip([1, 2, 3], [3, 2, 1]) if a != b]
