# vybe-test: python/while_loops/while_enumerate_manual_index_value
# origin: languages/python/tests/python/test_while_loops.rs

xs = ['x', 'y']
i = 0
while i < len(xs):
 print(str(i) + xs[i])
 i += 1
