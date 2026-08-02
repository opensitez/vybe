# vybe-test: python/pattern_matching_spec/match_mapping_guard_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

match {'x': 10}:
    case {'x': x} if x > 5:
        pass
