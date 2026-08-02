# vybe-test: python/generator_extended/generator_expr_map
# origin: languages/python/tests/python/test_generator_extended.rs

print(list(map(lambda x: x + 1, (i for i in range(3)))))
