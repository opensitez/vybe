# vybe-test: python/python_numbers_abstract_hierarchy/test_numbers_abc_instantiation_raises
# origin: languages/python/tests/python/test_python_numbers_abstract_hierarchy.rs

import numbers

try:
    numbers.Number()
    print('no_error')
except TypeError:
    print('TypeError_raised')
