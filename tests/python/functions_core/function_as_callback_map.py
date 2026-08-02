# vybe-test: python/functions_core/function_as_callback_map
# origin: languages/python/tests/python/test_functions_core.rs

def inc(x):
 return x + 1
list(map(inc, [1, 2, 3]))
