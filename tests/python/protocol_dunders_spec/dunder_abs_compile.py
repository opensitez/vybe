# vybe-test: python/protocol_dunders_spec/dunder_abs_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs
# vybe-test-mode: compile

class Num:
    def __abs__(self):
        return 0
