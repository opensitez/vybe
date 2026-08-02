# vybe-test: python/classes_extended/super_call
# origin: languages/python/tests/python/test_classes_extended.rs
# vybe-test-mode: compile

class Dog(Animal):
    def __init__(self, name):
        super().__init__(name)
