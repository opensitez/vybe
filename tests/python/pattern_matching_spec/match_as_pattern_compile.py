# vybe-test: python/pattern_matching_spec/match_as_pattern_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

match [1, 2]:
    case [1, x] as whole:
        print(x, whole)
