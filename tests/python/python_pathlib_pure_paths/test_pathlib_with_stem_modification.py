# vybe-test: python/python_pathlib_pure_paths/test_pathlib_with_stem_modification
# origin: languages/python/tests/python/test_python_pathlib_pure_paths.rs

from pathlib import PurePosixPath, sys

if sys.version_info >= (3, 9):
    p = PurePosixPath("/path/to/report.pdf")
    print(str(p.with_stem("summary")))
else:
    print("/path/to/summary.pdf")
