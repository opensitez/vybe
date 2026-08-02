# vybe-test: python/protocol_dunders_spec/descriptor_set_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class D:
    def __set__(self, obj, value):
        pass
