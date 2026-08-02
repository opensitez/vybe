# vybe-test: python/pattern_matching_spec/match_mapping_exact_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

match {'x': 1, 'y': 2}:
    case {'x': 1, 'y': y}:
        print(y)
