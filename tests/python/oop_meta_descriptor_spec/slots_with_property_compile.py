# vybe-test: python/oop_meta_descriptor_spec/slots_with_property_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs

class C:
    __slots__ = ('_x',)
    @property
    def x(self):
        return self._x
