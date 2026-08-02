# vybe-test: python/oop_meta_descriptor_spec/property_deleter_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class C:
    def __init__(self):
        self._x = 0
    @property
    def x(self):
        return self._x
    @x.deleter
    def x(self):
        del self._x
