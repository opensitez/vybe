# vybe-test: python/pattern_matching_spec/match_mapping_double_star_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

match {'a': 1, 'b': 2}:
    case {'a': a, **rest}:
        print(a, rest)
