# vybe-test: python/oop_meta_descriptor_spec/descriptor_set_name_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs

class D:
    def __set_name__(self, owner, name):
        self.name = name
class C:
    value = D()
