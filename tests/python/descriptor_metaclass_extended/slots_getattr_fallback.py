# vybe-test: python/descriptor_metaclass_extended/slots_getattr_fallback
# origin: languages/python/tests/python/test_descriptor_metaclass_extended.rs

class C:
 __slots__ = ('a',)
 def __init__(self):
  self.a = 1
try:
 C().b
 print('ok')
except AttributeError:
 print('err')
