# vybe-test: python/walrus_core/walrus_dict_comp
# origin: languages/python/tests/python/test_walrus_core.rs

{k: (v := k * 2) for k in range(2)}
