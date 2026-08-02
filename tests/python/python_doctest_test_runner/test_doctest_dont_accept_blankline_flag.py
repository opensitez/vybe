# vybe-test: python/python_doctest_test_runner/test_doctest_dont_accept_blankline_flag
# origin: languages/python/tests/python/test_python_doctest_test_runner.rs

import doctest

def blank_output():
    """
    >>> blank_output()
    <BLANKLINE>
    """
    print()

res = doctest.testmod()
print(res.failed)
