# vybe-test: python/while_loops/while_reads_dict_keys_via_popitem
# origin: languages/python/tests/python/test_while_loops.rs

d = {'a': 1, 'b': 2}
keys = []
while d:
 k, v = d.popitem()
 keys.append(k)
print(len(keys))
