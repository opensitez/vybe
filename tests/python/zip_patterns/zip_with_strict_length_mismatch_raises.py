# vybe-test: python/zip_patterns/zip_with_strict_length_mismatch_raises
# origin: languages/python/tests/python/test_zip_patterns.rs

try:
 list(zip([1, 2], [1], strict=True))
 print('ok')
except ValueError:
 print('ValueError')
