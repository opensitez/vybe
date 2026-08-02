# vybe-test: python/enumerate_zip_core/zip_sum_columns
# origin: languages/python/tests/python/test_enumerate_zip_core.rs

cols = list(zip([1, 2], [3, 4]))
print(sum(a + b for a, b in cols))
