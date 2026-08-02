# vybe-test: python/dict_pop_views/dict_items_unpack_in_for_loop
# origin: languages/python/tests/python/test_dict_pop_views.rs

total = 0
for k, v in {'a': 1, 'b': 2}.items():
 total += v
print(total)
