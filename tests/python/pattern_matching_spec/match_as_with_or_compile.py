# vybe-test: python/pattern_matching_spec/match_as_with_or_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs

match 1:
    case (1 | 2) as value:
        print(value)
