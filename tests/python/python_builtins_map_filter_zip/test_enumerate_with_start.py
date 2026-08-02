# vybe-test: python/python_builtins_map_filter_zip/test_enumerate_with_start
# origin: languages/python/tests/python/test_python_builtins_map_filter_zip.rs

for i, v in enumerate(['a', 'b', 'c'], start=5):
    print(i, v)
