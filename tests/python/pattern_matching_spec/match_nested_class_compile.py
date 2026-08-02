# vybe-test: python/pattern_matching_spec/match_nested_class_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

class Node:
    pass
match Node():
    case Node(left=Node(), right=_):
        pass
