# vybe-test: python/classes_extended/property_compile
# origin: languages/python/tests/python/test_classes_extended.rs
# vybe-test-mode: compile

class Circle:
    def __init__(self, r):
        self._r = r
    @property
    def radius(self):
        return self._r
