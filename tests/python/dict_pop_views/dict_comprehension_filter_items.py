# vybe-test: python/dict_pop_views/dict_comprehension_filter_items
# origin: languages/python/tests/python/test_dict_pop_views.rs

{k: v for k, v in {'a': 1, 'b': 2, 'c': 3}.items() if v > 1}
