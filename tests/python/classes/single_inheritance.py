# vybe-test: python/classes/single_inheritance
# origin: languages/python/tests/python/test_classes.rs
# vybe-test-mode: compile

class Animal:
    def speak(self):
        return 'generic'

class Dog(Animal):
    def speak(self):
        return 'woof'
