# vybe-test: python/descriptor_metaclass_extended/descriptor_getset
# origin: languages/python/tests/python/test_descriptor_metaclass_extended.rs

class D:
 def __get__(self, obj, owner): return 1
 def __set__(self, obj, val): pass
class C:
 x = D()
