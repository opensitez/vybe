# vybe-test: python/while_loops/while_list_comprehension_alternative_sum
# origin: languages/python/tests/python/test_while_loops.rs

xs = [1, 2, 3]
i = 0
total = 0
while i < len(xs):
 total += xs[i]
 i += 1
print(total)
