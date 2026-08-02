# vybe-test: python/walrus_extended/walrus_sum_comp
# origin: languages/python/tests/python/test_walrus_extended.rs

print(sum((v := i) for i in range(4)))
