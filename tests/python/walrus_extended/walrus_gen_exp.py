# vybe-test: python/walrus_extended/walrus_gen_exp
# origin: languages/python/tests/python/test_walrus_extended.rs

print(list((y := x) for x in range(2)))
