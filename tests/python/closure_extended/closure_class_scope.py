# vybe-test: python/closure_extended/closure_class_scope
# origin: languages/python/tests/python/test_closure_extended.rs
# vybe-test-mode: compile

def outer():
 class C:
  def m(self):
   return x
