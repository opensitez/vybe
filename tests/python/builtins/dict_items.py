# vybe-test: python/builtins/dict_items
# origin: languages/python/tests/python/test_builtins.rs
d = {'key': 1, 'a': 1}

for k, v in d.items():
    print(k, v)
