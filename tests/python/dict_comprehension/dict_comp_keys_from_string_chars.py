# vybe-test: python/dict_comprehension/dict_comp_keys_from_string_chars
# origin: languages/python/tests/python/test_dict_comprehension.rs

{c: ord(c) for c in 'ab' if c == 'a'}
