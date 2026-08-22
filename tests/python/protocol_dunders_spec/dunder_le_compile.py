# vybe-test: python/protocol_dunders_spec/dunder_le_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Box:
    def __le__(self, other):
        return True
