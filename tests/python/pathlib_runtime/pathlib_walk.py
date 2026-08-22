# vybe-test: python/pathlib_runtime/pathlib_walk
# origin: languages/python/tests/python/test_pathlib_runtime.rs

from pathlib import Path
[p for p in Path('.').iterdir()]
