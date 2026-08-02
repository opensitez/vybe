# vybe-test: python/pattern_matching_spec/match_subject_call_compile
# origin: languages/python/tests/python/test_pattern_matching_spec.rs
# vybe-test-mode: compile

def make():
    return [1, 2]
match make():
    case [1, x]:
        print(x)
