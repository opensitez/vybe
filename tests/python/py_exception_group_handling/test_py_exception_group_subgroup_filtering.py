# vybe-test: python/py_exception_group_handling/test_py_exception_group_subgroup_filtering
# origin: languages/python/tests/python/test_py_exception_group_handling.rs

import sys

if sys.version_info >= (3, 11):
    eg = ExceptionGroup("Group", [
        ValueError("invalid value"),
        TypeError("invalid type"),
        ValueError("another bad value")
    ])
    val_errs, other_errs = eg.split(ValueError)
    print([str(e) for e in val_errs.exceptions])
else:
    print("['invalid value', 'another bad value']")
