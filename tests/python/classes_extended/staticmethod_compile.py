# vybe-test: python/classes_extended/staticmethod_compile
# origin: languages/python/tests/python/test_classes_extended.rs
# vybe-test-mode: compile

class Math:
    @staticmethod
    def add(a, b):
        return a + b
