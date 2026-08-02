# vybe-test: python/pattern_matching_spec/match_tuple_or_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

match (1, 2):
    case (1, 2) | (2, 1):
        pass
