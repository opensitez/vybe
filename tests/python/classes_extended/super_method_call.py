# vybe-test: python/classes_extended/super_method_call
# origin: languages/python/tests/python/test_classes_extended.rs
# The base/name this fixture uses was never defined — supplied so it RUNS.
class Parent:
    def method(self):
        return 1


class Child(Parent):
    def method(self):
        return super().method() + 1
