# vybe-test: python/pattern_matching_spec/match_class_keyword_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs

class Point:
    pass
match Point():
    case Point(x=1, y=y):
        print(y)
