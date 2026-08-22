# vybe-test: python/protocol_dunders_spec/dunder_mul_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Vec:
    def __mul__(self, other):
        return self
