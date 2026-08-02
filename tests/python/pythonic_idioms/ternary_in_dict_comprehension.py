# vybe-test: python/pythonic_idioms/ternary_in_dict_comprehension
# origin: languages/python/tests/python/test_pythonic_idioms.rs

{i: ('pos' if i > 0 else 'nonpos') for i in [-1, 0, 1]}
