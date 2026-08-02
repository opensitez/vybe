# vybe-test: python/functions/map_with_named_func
# origin: languages/python/tests/python/test_functions.rs
# vybe-test-mode: compile

def double(x):
    return x * 2
result = map(double, [1, 2, 3])
