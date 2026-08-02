# vybe-test: python/classes/class_with_init
# origin: languages/python/tests/python/test_classes.rs
# vybe-test-mode: compile

class Dog:
    def __init__(self, name):
        self.name = name
    def bark(self):
        print(self.name)
