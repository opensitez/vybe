# vybe-test: python/function_signatures_extended/too_many_positional
# origin: languages/python/tests/python/test_function_signatures_extended.rs

def f(a, b):
 return a + b
try:
 f(1, 2, 3)
 print('ok')
except TypeError:
 print('err')
