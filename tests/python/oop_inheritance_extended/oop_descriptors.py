# vybe-test: python/oop_inheritance_extended/oop_descriptors
# origin: languages/python/tests/python/test_oop_inheritance_extended.rs
# vybe-test-mode: compile

class D:
 def __get__(self, obj, owner):
  return 1
class C:
 x = D()
