# vybe-test: python/walrus_core/walrus_list_comp_value_reuse
# origin: languages/python/tests/python/test_walrus_core.rs

[(s := str(i)) for i in range(2)]
