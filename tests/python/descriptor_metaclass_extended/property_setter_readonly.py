# vybe-test: python/descriptor_metaclass_extended/property_setter_readonly
# origin: languages/python/tests/python/test_descriptor_metaclass_extended.rs

class C:
 @property
 def x(self):
  return 1
try:
 C().x = 2
 print('ok')
except AttributeError:
 print('err')
