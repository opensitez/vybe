# vybe-test: python/classes_extended/super_method_call
# origin: languages/python/tests/python/test_classes_extended.rs
# vybe-test-mode: compile

class Child(Parent):
    def method(self):
        return super().method() + 1
