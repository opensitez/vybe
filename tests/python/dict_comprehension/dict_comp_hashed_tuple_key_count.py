# vybe-test: python/dict_comprehension/dict_comp_hashed_tuple_key_count
# origin: languages/python/tests/python/test_dict_comprehension.rs

d = {(a, b): a + b for a in [1] for b in [2, 3]}
print(len(d))
