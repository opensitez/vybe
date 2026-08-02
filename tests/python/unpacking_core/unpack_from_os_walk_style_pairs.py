# vybe-test: python/unpacking_core/unpack_from_os_walk_style_pairs
# origin: languages/python/tests/python/test_unpacking_core.rs

pairs = [('a', 1), ('b', 2)]
out = []
for k, v in pairs:
 out.append(k + str(v))
print(out)
