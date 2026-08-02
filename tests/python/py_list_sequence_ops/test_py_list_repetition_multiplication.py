# vybe-test: python/py_list_sequence_ops/test_py_list_repetition_multiplication
# origin: languages/python/tests/python/test_py_list_sequence_ops.rs

zeros = [0] * 5
print(zeros)

nested = [[]] * 3  # note: shared reference
nested[0].append(1)
print(nested)

# un-shared initialization pattern
unshared = [[] for _ in range(3)]
unshared[0].append(1)
print(unshared)
