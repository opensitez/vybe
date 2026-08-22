# vybe-test: python/new_features/class_tostring_alias
# origin: languages/python/tests/python/test_new_features.rs

class Dog:
    def __init__(self, name):
        self.name = name
    def __str__(self):
        return self.name
d = Dog('Rex')
