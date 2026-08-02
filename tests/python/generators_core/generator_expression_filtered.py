# vybe-test: python/generators_core/generator_expression_filtered
# origin: languages/python/tests/python/test_generators_core.rs

list(x for x in range(5) if x % 2 == 0)
