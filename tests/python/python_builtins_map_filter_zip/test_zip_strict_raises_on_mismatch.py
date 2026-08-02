# vybe-test: python/python_builtins_map_filter_zip/test_zip_strict_raises_on_mismatch
# origin: languages/python/tests/python/test_python_builtins_map_filter_zip.rs

try:
    list(zip([1, 2, 3], [1, 2], strict=True))
    print("no_error")
except ValueError:
    print("ValueError")
