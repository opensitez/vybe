# vybe-test: python/oop_meta_descriptor_spec/annotations_dict_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs
# vybe-test-mode: compile

class C:
    x: int
    y: str
ann = C.__annotations__
