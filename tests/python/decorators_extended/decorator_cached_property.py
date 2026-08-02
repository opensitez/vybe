# vybe-test: python/decorators_extended/decorator_cached_property
# origin: languages/python/tests/python/test_decorators_extended.rs
# vybe-test-mode: compile

class C:
 @property
 def x(self):
  return 1
