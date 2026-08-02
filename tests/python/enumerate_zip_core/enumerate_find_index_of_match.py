# vybe-test: python/enumerate_zip_core/enumerate_find_index_of_match
# origin: languages/python/tests/python/test_enumerate_zip_core.rs

target = 'y'
idx = next(i for i, v in enumerate(['x', 'y']) if v == target)
print(idx)
