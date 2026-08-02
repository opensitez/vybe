# vybe-test: python/closure_extended/closure_nonlocal_error
# origin: languages/python/tests/python/test_closure_extended.rs
# vybe-test-mode: compile

def outer():
 def inner():
  nonlocal x
