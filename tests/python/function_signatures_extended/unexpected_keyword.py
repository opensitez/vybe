# vybe-test: python/function_signatures_extended/unexpected_keyword
# origin: languages/python/tests/python/test_function_signatures_extended.rs

def f(a):
 return a
try:
 f(a=1, b=2)
 print('ok')
except TypeError:
 print('err')
