# vybe-test: python/functions_core/function_returns_list_mutation_visible
# origin: languages/python/tests/python/test_functions_core.rs

def f():
 return [1]
x = f()
x.append(2)
x
