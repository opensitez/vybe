# vybe-test: python/new_features/match_basic_value
# origin: languages/python/tests/python/test_new_features.rs

x = 42
match x:
    case 1:
        print('one')
    case 42:
        print('forty-two')
    case _:
        print('other')
