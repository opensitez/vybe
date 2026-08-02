# vybe-test: python/enumerate_zip_core/enumerate_count_filtered_items
# origin: languages/python/tests/python/test_enumerate_zip_core.rs

print(sum(1 for i, v in enumerate([1, 2, 3]) if v > 1))
