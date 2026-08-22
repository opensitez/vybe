# vybe-test: python/oop_inheritance_extended/oop_init_subclass
# origin: languages/python/tests/python/test_oop_inheritance_extended.rs

class B:
 def __init_subclass__(cls, **kw):
  pass
class D(B):
 pass
