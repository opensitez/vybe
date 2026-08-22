# vybe-test: python/classes/class_property_decorator
# origin: languages/python/tests/python/test_classes.rs

class Circle:
    def __init__(self, r):
        self.r = r
    @property
    def radius(self):
        return self.r
