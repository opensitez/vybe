# vybe-test: python/walrus_extended/walrus_dict_comp
# origin: languages/python/tests/python/test_walrus_extended.rs

print({k: (v := k * 2) for k in range(3)})
