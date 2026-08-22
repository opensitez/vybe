# vybe-test: python/oop_meta_descriptor_spec/class_getitem_compile
# origin: languages/python/tests/python/test_oop_meta_descriptor_spec.rs

class Box:
    def __class_getitem__(cls, item):
        return cls
T = Box[int]
