# vybe-test: python/closure_extended/closure_annotations
# origin: languages/python/tests/python/test_closure_extended.rs
# vybe-test-mode: compile

def outer():
 x: int = 1
 def inner() -> int:
  return x
