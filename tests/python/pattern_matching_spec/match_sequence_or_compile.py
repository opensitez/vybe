# vybe-test: python/pattern_matching_spec/match_sequence_or_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs

match [1, 2]:
    case [1, 2] | [2, 1]:
        pass
