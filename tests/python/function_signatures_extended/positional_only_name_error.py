# vybe-test: python/function_signatures_extended/positional_only_name_error
# origin: languages/python/tests/python/test_function_signatures_extended.rs

def f(a, /, b):
 return a + b
try:
 f(a=1, b=2)
 print('ok')
except TypeError:
 print('err')
