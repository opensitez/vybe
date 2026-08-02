# vybe-test: python/set_core/set_iter_order_insertion
# origin: languages/python/tests/python/test_set_core.rs

s = set()
for x in [3, 1, 2]:
 s.add(x)
print(list(s))
