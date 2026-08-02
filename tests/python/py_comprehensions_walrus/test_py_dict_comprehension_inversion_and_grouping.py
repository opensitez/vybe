# vybe-test: python/py_comprehensions_walrus/test_py_dict_comprehension_inversion_and_grouping
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

words = ["apple", "ant", "bear", "bat", "cat"]
by_first = {
    letter: [w for w in words if w.startswith(letter)]
    for letter in set(w[0] for w in words)
}
print(sorted(by_first.keys()))
print(sorted(by_first["a"]))
print(sorted(by_first["b"]))
