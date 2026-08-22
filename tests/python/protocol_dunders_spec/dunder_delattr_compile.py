# vybe-test: python/protocol_dunders_spec/dunder_delattr_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Proxy:
    def __delattr__(self, name):
        super().__delattr__(name)
