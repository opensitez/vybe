# vybe-test: python/generators_core/generator_expression_in_sum
# origin: languages/python/tests/python/test_generators_core.rs

print(sum(x * x for x in range(4)))
