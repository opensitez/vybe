# vybe-test: python/protocol_dunders_spec/descriptor_delete_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class D:
    def __delete__(self, obj):
        pass
