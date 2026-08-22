# vybe-test: python/pattern_matching_spec/match_class_positional_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs

class Point:
    __match_args__ = ('x', 'y')
match Point():
    case Point(1, y):
        print(y)
