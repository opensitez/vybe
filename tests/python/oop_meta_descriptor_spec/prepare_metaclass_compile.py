# vybe-test: python/oop_meta_descriptor_spec/prepare_metaclass_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class Meta(type):
    @classmethod
    def __prepare__(mcls, name, bases):
        return {}
class C(metaclass=Meta):
    pass
