# vybe-test: python/protocol_dunders_spec/dunder_setattr_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Proxy:
    def __setattr__(self, name, value):
        super().__setattr__(name, value)
