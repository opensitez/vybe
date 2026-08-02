# vybe-test: python/while_loops/while_finds_first_negative
# origin: languages/python/tests/python/test_while_loops.rs

xs = [1, 3, -2, 4]
i = 0
while i < len(xs) and xs[i] >= 0:
 i += 1
print(xs[i])
