# vybe-test: python/python_extended_unpacking/test_unpack_in_for_loop
# origin: languages/python/tests/python/test_python_extended_unpacking.rs

pairs = [(1, 'a'), (2, 'b'), (3, 'c')]
for num, letter in pairs:
    print(num, letter)
