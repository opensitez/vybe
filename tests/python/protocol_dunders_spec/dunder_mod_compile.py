# vybe-test: python/protocol_dunders_spec/dunder_mod_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Vec:
    def __mod__(self, other):
        return self
