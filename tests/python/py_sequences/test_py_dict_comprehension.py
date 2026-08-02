# vybe-test: python/py_sequences/test_py_dict_comprehension
# origin: languages/python/tests/python/test_py_sequences.rs

squares = {x: x**2 for x in range(5)}
print(squares)

inverted = {v: k for k, v in squares.items()}
print(inverted)

filtered = {k: v for k, v in squares.items() if v > 4}
print(filtered)
