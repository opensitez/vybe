# vybe-test: python/enumerate_start_index/enumerate_unpack_in_loop
# origin: languages/python/tests/python/test_enumerate_start_index.rs

s = ''
for i, ch in enumerate('ab'):
 s += str(i) + ch
print(s)
