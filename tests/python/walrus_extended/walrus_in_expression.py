# vybe-test: python/walrus_extended/walrus_in_expression
# origin: languages/python/tests/python/test_walrus_extended.rs

print([(a := i) for i in range(2)][-1])
