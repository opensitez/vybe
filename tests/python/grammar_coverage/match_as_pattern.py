# vybe-test: python/grammar_coverage/match_as_pattern
# origin: languages/python/tests/python/test_grammar_coverage.rs

match [1, 2]:
    case [x, y] as point:
        print(point)
    case _:
        pass
