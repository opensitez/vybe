# vybe-test: python/while_loops/while_truthy_empty_list_stops
# origin: languages/python/tests/python/test_while_loops.rs

xs = [1]
while xs:
 xs.pop()
print('end')
