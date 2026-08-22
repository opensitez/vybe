# vybe-test: python/new_features/match_or_pattern
# origin: languages/python/tests/python/test_new_features.rs

x = 2
match x:
    case 1 | 2 | 3:
        print('small')
    case _:
        print('big')
