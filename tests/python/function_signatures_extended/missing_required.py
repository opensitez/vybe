# vybe-test: python/function_signatures_extended/missing_required
# origin: languages/python/tests/python/test_function_signatures_extended.rs

def f(a, b):
 return a + b
try:
 f(1)
 print('ok')
except TypeError:
 print('err')
