# vybe-test: python/py_sequences/test_py_list_comprehension_basic_and_filtered
# origin: languages/python/tests/python/test_py_sequences.rs

squares = [x ** 2 for x in range(6)]
print(squares)

evens = [x for x in range(10) if x % 2 == 0]
print(evens)

nested = [[i * j for j in range(1, 4)] for i in range(1, 4)]
print(nested)
