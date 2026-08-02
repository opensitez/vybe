# vybe-test: python/pattern_matching_spec/match_wildcard_after_sequence_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

match [1, 2, 3]:
    case [1, *_]:
        pass
