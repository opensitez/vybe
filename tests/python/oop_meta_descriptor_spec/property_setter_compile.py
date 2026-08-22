# vybe-test: python/oop_meta_descriptor_spec/property_setter_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs

class C:
    def __init__(self):
        self._x = 0
    @property
    def x(self):
        return self._x
    @x.setter
    def x(self, value):
        self._x = value
