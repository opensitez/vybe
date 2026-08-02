# vybe-test: python/oop_meta_descriptor_spec/instance_dict_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class C:
    pass
c = C()
d = c.__dict__
