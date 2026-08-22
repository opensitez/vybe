# vybe-test: python/protocol_dunders_spec/dunder_pos_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Num:
    def __pos__(self):
        return self
