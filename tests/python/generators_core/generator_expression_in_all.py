# vybe-test: python/generators_core/generator_expression_in_all
# origin: languages/python/tests/python/test_generators_core.rs

print(all(x > 0 for x in [1, 2]))
