# vybe-test: python/dict_methods_extended/dict_popitem_empty_raises_compile
# origin: languages/python/tests/python/test_dict_methods_extended.rs

d = {}
try:
    d.popitem()
    print('ok')
except KeyError:
    print('err')
