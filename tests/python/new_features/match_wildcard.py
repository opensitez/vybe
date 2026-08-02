# vybe-test: python/new_features/match_wildcard
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

match 'hello':
    case _:
        print('anything')
