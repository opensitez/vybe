# vybe-test: python/protocol_dunders_spec/dunder_getattribute_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Proxy:
    def __getattribute__(self, name):
        return super().__getattribute__(name)
