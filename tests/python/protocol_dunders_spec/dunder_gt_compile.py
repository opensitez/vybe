# vybe-test: python/protocol_dunders_spec/dunder_gt_compile
# origin: languages/python/tests/python/test_protocol_dunders_spec.rs

class Box:
    def __gt__(self, other):
        return False
