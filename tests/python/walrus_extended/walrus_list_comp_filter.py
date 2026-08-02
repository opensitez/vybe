# vybe-test: python/walrus_extended/walrus_list_comp_filter
# origin: languages/python/tests/python/test_walrus_extended.rs

print([y for x in [1, 2, 3] if (y := x * 2) > 2])
