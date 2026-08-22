# vybe-test: python/protocol_dunders_spec/dunder_truediv_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Vec:
    def __truediv__(self, other):
        return self
