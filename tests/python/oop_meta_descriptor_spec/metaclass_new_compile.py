# vybe-test: python/oop_meta_descriptor_spec/metaclass_new_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs

class Meta(type):
    def __new__(mcls, name, bases, ns):
        return super().__new__(mcls, name, bases, ns)
