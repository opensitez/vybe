# vybe-test: python/classes_extended/super_call
# origin: languages/python/tests/python/test_classes_extended.rs
# The base/name this fixture uses was never defined — supplied so it RUNS.
class Animal:
    def __init__(self, name):
        self.name = name


class Dog(Animal):
    def __init__(self, name):
        super().__init__(name)
