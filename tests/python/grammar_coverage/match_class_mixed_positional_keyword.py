# vybe-test: python/grammar_coverage/match_class_mixed_positional_keyword
# origin: languages/python/tests/python/test_grammar_coverage.rs
# vybe-test-mode: compile

match p:
    case Rect(a, width=10):
        pass
    case _:
        pass
