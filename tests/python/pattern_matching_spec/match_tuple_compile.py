# vybe-test: python/pattern_matching_spec/match_tuple_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

match (1, 2):
    case (1, x):
        print(x)
