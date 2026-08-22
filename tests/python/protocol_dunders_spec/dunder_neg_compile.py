# vybe-test: python/protocol_dunders_spec/dunder_neg_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Num:
    def __neg__(self):
        return self
