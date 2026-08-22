# vybe-test: python/pattern_matching_spec/match_list_of_tuples_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs

match [('a', 1), ('b', 2)]:
    case [('a', x), *rest]:
        print(x, rest)
