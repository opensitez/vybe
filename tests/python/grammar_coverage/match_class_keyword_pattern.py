# vybe-test: python/grammar_coverage/match_class_keyword_pattern
# origin: languages/python/tests/python/test_grammar_coverage.rs
# vybe-test-mode: compile

class Point:
    x = 0
    y = 0
match p:
    case Point(x=1, y=2):
        print('origin-ish')
    case _:
        pass
