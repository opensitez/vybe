# vybe-test: python/oop_meta_descriptor_spec/datamodel_repr_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class C:
    def __repr__(self):
        return 'C()'
    def __str__(self):
        return 'c'
