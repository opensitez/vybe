# vybe-test: python/pythonic_idioms/generator_expr_in_all
# origin: languages/python/tests/python/test_pythonic_idioms.rs

print(all(x > 0 for x in [1, 2, 3]))
