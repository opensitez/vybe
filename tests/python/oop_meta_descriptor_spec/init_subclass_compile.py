# vybe-test: python/oop_meta_descriptor_spec/init_subclass_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class Base:
    def __init_subclass__(cls, flag=False, **kwargs):
        super().__init_subclass__(**kwargs)
class Child(Base, flag=True):
    pass
