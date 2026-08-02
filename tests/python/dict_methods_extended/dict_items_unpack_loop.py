# vybe-test: python/dict_methods_extended/dict_items_unpack_loop
# origin: languages/python/tests/python/test_dict_methods_extended.rs

d = {'x': 10, 'y': 20}
s = 0
for k, v in d.items():
    s += v
print(s)
