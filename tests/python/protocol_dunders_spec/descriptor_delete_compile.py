# vybe-test: python/protocol_dunders_spec/descriptor_delete_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class D:
    def __delete__(self, obj):
        pass
