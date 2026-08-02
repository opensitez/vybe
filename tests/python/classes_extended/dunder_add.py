# vybe-test: python/classes_extended/dunder_add
# origin: languages/python/tests/python/test_classes_extended.rs
# vybe-test-mode: compile

class Vec:
    def __init__(self, x, y):
        self.x = x
        self.y = y
    def __add__(self, other):
        return Vec(self.x + other.x, self.y + other.y)
