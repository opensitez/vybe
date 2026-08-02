# vybe-test: python/dict_core/dict_items_unpack_in_for
# origin: languages/python/tests/python/test_dict_core.rs

total = 0
for k, v in {'a': 1, 'b': 2}.items():
 total += v
print(total)
