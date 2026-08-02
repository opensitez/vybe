# vybe-test: python/classes/staticmethod_no_self_param
# origin: languages/python/tests/python/test_classes.rs
# vybe-test-mode: compile

class Config:
    @staticmethod
    def default_value():
        return 42
v = Config.default_value()
