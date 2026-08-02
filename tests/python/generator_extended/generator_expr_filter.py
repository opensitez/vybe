# vybe-test: python/generator_extended/generator_expr_filter
# origin: languages/python/tests/python/test_generator_extended.rs

print(list(x for x in range(5) if x % 2))
