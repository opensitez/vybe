# vybe-test: python/enumerate_zip_core/enumerate_string_find_position
# origin: languages/python/tests/python/test_enumerate_zip_core.rs

s = 'banana'
print(next(i for i, ch in enumerate(s) if ch == 'n'))
