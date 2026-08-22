# vybe-test: python/decorators_extended/decorator_classmethod_property
# origin: languages/python/tests/python/test_decorators_extended.rs

class C:
 @classmethod
 @property
 def x(cls):
  return 1
