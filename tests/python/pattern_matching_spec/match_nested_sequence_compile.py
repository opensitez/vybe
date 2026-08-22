# vybe-test: python/pattern_matching_spec/match_nested_sequence_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs

match [1, [2, 3]]:
    case [1, [a, b]]:
        print(a, b)
