# vybe-test: python/python_numbers_abstract_hierarchy/test_numbers_abc_dir_and_all
# origin: languages/python/tests/python/test_python_numbers_abstract_hierarchy.rs

import numbers

names = dir(numbers)
for expected in ['Number', 'Complex', 'Real', 'Rational', 'Integral']:
    print(expected in names)
