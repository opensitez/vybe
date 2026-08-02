# vybe-test: python/zip_patterns/zip_sorted_unique_keys_with_values
# origin: languages/python/tests/python/test_zip_patterns.rs

d = {'b': 2, 'a': 1}
print(list(zip(sorted(d), [d[k] for k in sorted(d)])))
