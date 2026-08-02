# vybe-test: python/generators_core/generator_expression_in_any
# origin: languages/python/tests/python/test_generators_core.rs

print(any(x > 2 for x in [1, 2, 3]))
