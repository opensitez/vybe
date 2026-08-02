# vybe-test: python/python_pathlib_pure_paths/test_pathlib_is_relative_to_check
# origin: languages/python/tests/python/test_python_pathlib_pure_paths.rs

from pathlib import PurePosixPath, sys

if sys.version_info >= (3, 9):
    p = PurePosixPath("/var/log/syslog")
    print(p.is_relative_to("/var/log"))
    print(p.is_relative_to("/etc"))
else:
    print("True\nFalse")
