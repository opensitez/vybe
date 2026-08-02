# vybe-test: python/classes_extended/class_with_metaclass
# origin: languages/python/tests/python/test_classes_extended.rs
# vybe-test-mode: compile

class Meta(type):
    pass
class MyClass(metaclass=Meta):
    pass
