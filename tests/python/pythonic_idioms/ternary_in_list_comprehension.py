# vybe-test: python/pythonic_idioms/ternary_in_list_comprehension
# origin: languages/python/tests/python/test_pythonic_idioms.rs

[('even' if i % 2 == 0 else 'odd') for i in range(3)]
