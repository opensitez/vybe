# vybe-test: python/new_features/match_none
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

x = None
match x:
    case None:
        print('none')
    case _:
        print('other')
