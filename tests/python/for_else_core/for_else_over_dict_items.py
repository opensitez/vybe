# vybe-test: python/for_else_core/for_else_over_dict_items
# origin: languages/python/tests/python/test_for_else_core.rs

total = 0
for k, v in {'a': 1, 'b': 2}.items():
 total += v
else:
 print(total)
