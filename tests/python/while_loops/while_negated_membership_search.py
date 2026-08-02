# vybe-test: python/while_loops/while_negated_membership_search
# origin: languages/python/tests/python/test_while_loops.rs

xs = [1, 2, 3]
target = 4
i = 0
while i < len(xs) and xs[i] != target:
 i += 1
print(i == len(xs))
