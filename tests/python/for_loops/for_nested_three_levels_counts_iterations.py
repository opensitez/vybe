# vybe-test: python/for_loops/for_nested_three_levels_counts_iterations
# origin: languages/python/tests/python/test_for_loops.rs

count = 0
for a in range(2):
    for b in range(2):
        for c in range(2):
            count += 1
print(count)
