# vybe-test: python/for_loops/for_dict_items_filter_by_value
# origin: languages/python/tests/python/test_for_loops.rs

picked = 0
for key, value in {'a': 1, 'b': 10, 'c': 3}.items():
    if value > 5:
        picked += 1
print(picked)
