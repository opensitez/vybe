# vybe-test: python/if_else/truthiness_nonempty_dict_is_truthy
# origin: languages/python/tests/python/test_if_else.rs

if {'a': 1}:
    print('yes')
else:
    print('no')
