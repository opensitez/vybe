# vybe-test: python/unpacking_core/unpack_matrix_rows
# origin: languages/python/tests/python/test_unpacking_core.rs

rows = [(1, 2), (3, 4)]
s = 0
for x, y in rows:
 s += x + y
print(s)
