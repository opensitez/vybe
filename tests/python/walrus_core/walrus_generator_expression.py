# vybe-test: python/walrus_core/walrus_generator_expression
# origin: languages/python/tests/python/test_walrus_core.rs

print(sum((n := x) for x in range(3)))
