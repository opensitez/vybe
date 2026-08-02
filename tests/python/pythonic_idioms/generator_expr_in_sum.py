# vybe-test: python/pythonic_idioms/generator_expr_in_sum
# origin: languages/python/tests/python/test_pythonic_idioms.rs

print(sum(x * x for x in range(4)))
