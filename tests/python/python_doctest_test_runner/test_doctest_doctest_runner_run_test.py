# vybe-test: python/python_doctest_test_runner/test_doctest_doctest_runner_run_test
# origin: languages/python/tests/python/test_python_doctest_test_runner.rs

import doctest

def greet(name):
    """
    >>> greet("World")
    'Hello World'
    """
    return f"Hello {name}"

finder = doctest.DocTestFinder()
runner = doctest.DocTestRunner(verbose=False)
tests = finder.find(greet)
for t in tests:
    res = runner.run(t)

print(runner.failures)
print(runner.tries)
