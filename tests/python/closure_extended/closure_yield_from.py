# vybe-test: python/closure_extended/closure_yield_from
# origin: languages/python/tests/python/test_closure_extended.rs

def outer():
 def inner():
  yield from range(2)
 return inner
