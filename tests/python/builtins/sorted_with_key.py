# vybe-test: python/builtins/sorted_with_key
# origin: languages/python/tests/python/test_builtins.rs

words = ['banana', 'apple', 'cherry']
result = sorted(words, key=len)
