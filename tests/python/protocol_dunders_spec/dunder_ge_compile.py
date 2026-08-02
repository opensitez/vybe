# vybe-test: python/protocol_dunders_spec/dunder_ge_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Box:
    def __ge__(self, other):
        return False
