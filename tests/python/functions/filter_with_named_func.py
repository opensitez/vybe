# vybe-test: python/functions/filter_with_named_func
# origin: languages/python/tests/python/test_functions.rs

def is_even(x):
    return x % 2 == 0
result = filter(is_even, [1, 2, 3, 4])
