# vybe-test: python/enumerate_zip_core/map_with_zip_unpack
# origin: languages/python/tests/python/test_enumerate_zip_core.rs

list(map(lambda p: p[0] + len(p[1]), zip([1, 2], ['a', 'bb'])))
