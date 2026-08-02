# vybe-test: python/python_pathlib_pure_paths/test_pathlib_parents_sequence
# origin: languages/python/tests/python/test_python_pathlib_pure_paths.rs

from pathlib import PurePosixPath
p = PurePosixPath("/a/b/c/d")
print([str(parent) for parent in p.parents])
