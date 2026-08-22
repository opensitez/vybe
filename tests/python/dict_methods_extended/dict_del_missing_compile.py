# vybe-test: python/dict_methods_extended/dict_del_missing_compile
# origin: languages/python/tests/python/test_dict_methods_extended.rs

d = {}
try:
    del d['x']
except KeyError:
    pass
