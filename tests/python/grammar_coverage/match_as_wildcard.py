# vybe-test: python/grammar_coverage/match_as_wildcard
# origin: languages/python/tests/python/test_grammar_coverage.rs

match 42:
    case x as val:
        print(val)
