# vybe-test: python/zip_patterns/zip_longest_fillvalue
# origin: languages/python/tests/python/test_zip_patterns.rs

list(__import__('itertools').zip_longest([1, 2], [3], fillvalue=0))
