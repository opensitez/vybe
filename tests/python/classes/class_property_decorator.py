# vybe-test: python/classes/class_property_decorator
# origin: languages/python/tests/python/test_classes.rs
# vybe-test-mode: compile

class Circle:
    def __init__(self, r):
        self.r = r
    @property
    def radius(self):
        return self.r
