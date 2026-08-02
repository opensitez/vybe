# vybe-test: python/raise_assert/raise_unbound_local_error
# origin: languages/python/tests/python/test_raise_assert.rs

def f():
 print(x)
 x = 1
try:
 f()
except UnboundLocalError:
 print('unbound')
