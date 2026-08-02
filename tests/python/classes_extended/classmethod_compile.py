# vybe-test: python/classes_extended/classmethod_compile
# origin: languages/python/tests/python/test_classes_extended.rs
# vybe-test-mode: compile

class Factory:
    count = 0
    @classmethod
    def create(cls):
        cls.count += 1
        return Factory()
