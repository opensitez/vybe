# vybe-test: python/new_features/match_with_guard
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

x = 10
match x:
    case n if n > 5:
        print('big')
    case _:
        print('small')
