# vybe-test: python/pythonic_idioms/walrus_in_comprehension_filter
# origin: languages/python/tests/python/test_pythonic_idioms.rs

[y for x in [1, 2, 3] if (y := x * 2) > 2]
