# vybe-test: python/for_loops/for_list_counts_elements_matching_predicate
# origin: languages/python/tests/python/test_for_loops.rs

count = 0
for n in [1, 2, 3, 4, 5]:
    if n % 2 == 0:
        count += 1
print(count)
