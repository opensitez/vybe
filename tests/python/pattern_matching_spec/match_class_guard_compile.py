# vybe-test: python/pattern_matching_spec/match_class_guard_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

class Box:
    pass
match Box():
    case Box() as box if box is not None:
        pass
