# vybe-test: python/grammar_coverage/match_as_singleton
# origin: languages/python/tests/python/test_grammar_coverage.rs
# vybe-test-mode: compile

match True:
    case True as v:
        pass
    case _:
        pass
