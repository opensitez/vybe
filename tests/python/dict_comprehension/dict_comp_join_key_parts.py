# vybe-test: python/dict_comprehension/dict_comp_join_key_parts
# origin: languages/python/tests/python/test_dict_comprehension.rs

{'-'.join([str(a), str(b)]): a + b for a, b in [(1, 2)]}
