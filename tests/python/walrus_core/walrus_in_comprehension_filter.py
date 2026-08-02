# vybe-test: python/walrus_core/walrus_in_comprehension_filter
# origin: languages/python/tests/python/test_walrus_core.rs

[y for x in [1, 2, 3] if (y := x * 2) > 2]
