# vybe-test: python/new_features/match_string
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

cmd = 'quit'
match cmd:
    case 'quit' | 'exit':
        print('bye')
    case 'help':
        print('help')
    case _:
        pass
