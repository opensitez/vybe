# vybe-test: python/for_loops/for_dict_break_on_matching_key
# origin: languages/python/tests/python/test_for_loops.rs

for key in {'a': 1, 'stop': 2, 'c': 3}:
    if key == 'stop':
        break
    print(key)
