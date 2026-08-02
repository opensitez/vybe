# vybe-test: python/while_loops/while_waits_until_predicate_true
# origin: languages/python/tests/python/test_while_loops.rs

xs = [1, 2, 5]
i = 0
while i < len(xs) and xs[i] < 5:
 i += 1
print(xs[i])
