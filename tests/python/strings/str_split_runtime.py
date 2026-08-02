# vybe-test: python/strings/str_split_runtime
# origin: languages/python/tests/python/test_strings.rs

parts = 'a,b,c'.split(',')
for p in parts:
    print(p)
