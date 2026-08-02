# vybe-test: python/classes/classmethod_basic
# origin: languages/python/tests/python/test_classes.rs
# vybe-test-mode: compile

class Foo:
    @classmethod
    def create(cls):
        return Foo()
    def __init__(self):
        pass
