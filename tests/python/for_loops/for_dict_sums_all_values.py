# vybe-test: python/for_loops/for_dict_sums_all_values
# origin: languages/python/tests/python/test_for_loops.rs

total = 0
for v in {'a': 1, 'b': 2, 'c': 3}.values():
    total += v
print(total)
