# vybe-test: python/oop_meta_descriptor_spec/metaclass_call_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class Meta(type):
    def __call__(cls, *args, **kwargs):
        return super().__call__(*args, **kwargs)
