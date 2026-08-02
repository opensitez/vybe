# vybe-test: python/py_comprehensions_walrus/test_py_walrus_set_comprehension
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

words = ["hello", "world", "hi", "hey", "world"]
# Use walrus to capture length while deduplicating
long_words = {(word, length) for word in words if (length := len(word)) > 3}
print(sorted(long_words))
