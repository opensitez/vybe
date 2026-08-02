# vybe-test: python/zip_patterns/zip_parallel_accumulate_running_total
# origin: languages/python/tests/python/test_zip_patterns.rs

total = 0
for v, delta in zip([1, 2, 3], [1, 1, 1]):
 total += v * delta
print(total)
