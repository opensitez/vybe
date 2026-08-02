# vybe-test: python/closure_extended/closure_global_nonlocal_mix
# origin: languages/python/tests/python/test_closure_extended.rs
# vybe-test-mode: compile

x = 1
def outer():
 def inner():
  global x
  x = 2
