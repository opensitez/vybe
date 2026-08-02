# vybe-test: python/classes/multiple_inheritance
# origin: languages/python/tests/python/test_classes.rs
# vybe-test-mode: compile

class A:
    def method_a(self):
        return 'a'

class B:
    def method_b(self):
        return 'b'

class C(A, B):
    pass
