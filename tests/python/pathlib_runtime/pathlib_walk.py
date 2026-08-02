# vybe-test: python/pathlib_runtime/pathlib_walk
# origin: languages/python/tests/python/test_pathlib_runtime.rs
# vybe-test-mode: compile

from pathlib import Path
[p for p in Path('.').iterdir()]
