# vybe-test: python/python_warnings_categories/test_warnings_multiple_warnings_recorded
# origin: languages/python/tests/python/test_python_warnings_categories.rs

import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn("first", UserWarning)
    warnings.warn("second", DeprecationWarning)
    warnings.warn("third", RuntimeWarning)
print(len(w))
print([x.category.__name__ for x in w])
