# vybe-test: python/classes/class_with_inheritance
# origin: languages/python/tests/python/test_classes.rs
# vybe-test-mode: compile

class Animal:
    def speak(self):
        pass

class Dog(Animal):
    def bark(self):
        pass
