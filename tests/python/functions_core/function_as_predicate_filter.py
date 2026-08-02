# vybe-test: python/functions_core/function_as_predicate_filter
# origin: languages/python/tests/python/test_functions_core.rs

def is_pos(x):
 return x > 0
list(filter(is_pos, [-1, 0, 2]))
