# vybe-test: python/dict_comprehension/dict_comp_items_roundtrip
# origin: languages/python/tests/python/test_dict_comprehension.rs

d = {k: v * 2 for k, v in {'a': 1}.items()}
print(list(d.items())[0])
