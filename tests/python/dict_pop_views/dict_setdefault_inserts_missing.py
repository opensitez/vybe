# vybe-test: python/dict_pop_views/dict_setdefault_inserts_missing
# origin: languages/python/tests/python/test_dict_pop_views.rs

d = {}
print(d.setdefault('k', 5))
print(d['k'])
