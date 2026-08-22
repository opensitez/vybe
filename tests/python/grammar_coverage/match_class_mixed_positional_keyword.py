# vybe-test: python/grammar_coverage/match_class_mixed_positional_keyword
# origin: languages/python/tests/python/test_grammar_coverage.rs
# The base/name this fixture uses was never defined — supplied so it RUNS.
class Rect:
    __match_args__ = ('a',)
    def __init__(self, a, width):
        self.a = a
        self.width = width
p = Rect(1, 10)


match p:
    case Rect(a, width=10):
        pass
    case _:
        pass
