# vybe-test: python/for_loops/for_string_counts_specific_letter
# origin: languages/python/tests/python/test_for_loops.rs

count = 0
for ch in 'banana':
    if ch == 'a':
        count += 1
print(count)
