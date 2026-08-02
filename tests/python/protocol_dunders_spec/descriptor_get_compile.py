# vybe-test: python/protocol_dunders_spec/descriptor_get_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class D:
    def __get__(self, obj, owner):
        return 1
