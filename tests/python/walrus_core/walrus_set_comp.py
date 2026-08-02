# vybe-test: python/walrus_core/walrus_set_comp
# origin: languages/python/tests/python/test_walrus_core.rs

sorted({(y := x + 1) for x in range(2)})
