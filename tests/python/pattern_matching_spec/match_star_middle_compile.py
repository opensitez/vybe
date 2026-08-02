# vybe-test: python/pattern_matching_spec/match_star_middle_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

match [1, 2, 3, 4]:
    case [first, *middle, last]:
        print(first, middle, last)
