# vybe-test: python/classes_extended/class_multiple_bases
# origin: languages/python/tests/python/test_classes_extended.rs
# vybe-test-mode: compile

class A:
    pass
class B:
    pass
class C(A, B):
    pass
