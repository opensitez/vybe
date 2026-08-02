# vybe-test: python/function_signatures_extended/keyword_only_error
# origin: languages/python/tests/python/test_function_signatures_extended.rs

def f(a, *, b):
 return a + b
try:
 f(1, 2)
 print('ok')
except TypeError:
 print('err')
