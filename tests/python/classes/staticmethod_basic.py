# vybe-test: python/classes/staticmethod_basic
# origin: languages/python/tests/python/test_classes.rs
# vybe-test-mode: compile

class Math:
    @staticmethod
    def add(a, b):
        return a + b
result = Math.add(1, 2)
