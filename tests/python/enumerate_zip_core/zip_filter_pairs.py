# vybe-test: python/enumerate_zip_core/zip_filter_pairs
# origin: languages/python/tests/python/test_enumerate_zip_core.rs

[(a, b) for a, b in zip([1, 2, 3], [0, 2, 0]) if b]
