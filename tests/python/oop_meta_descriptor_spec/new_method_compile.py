# vybe-test: python/oop_meta_descriptor_spec/new_method_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class C:
    def __new__(cls, *args, **kwargs):
        return super().__new__(cls)
