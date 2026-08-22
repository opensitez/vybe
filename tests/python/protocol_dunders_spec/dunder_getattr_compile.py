# vybe-test: python/protocol_dunders_spec/dunder_getattr_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Proxy:
    def __getattr__(self, name):
        return 0
