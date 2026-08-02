# vybe-test: python/pathlib_runtime/pathlib_symlink
# origin: languages/python/tests/python/test_pathlib_runtime.rs
# vybe-test-mode: compile

from pathlib import Path
hasattr(Path('.'), 'symlink_to')
