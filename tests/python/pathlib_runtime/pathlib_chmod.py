# vybe-test: python/pathlib_runtime/pathlib_chmod
# origin: languages/python/tests/python/test_pathlib_runtime.rs
# vybe-test-mode: compile

from pathlib import Path
hasattr(Path('.'), 'chmod')
