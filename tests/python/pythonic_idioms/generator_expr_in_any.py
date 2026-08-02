# vybe-test: python/pythonic_idioms/generator_expr_in_any
# origin: languages/python/tests/python/test_pythonic_idioms.rs

print(any(x > 2 for x in [1, 2, 3]))
