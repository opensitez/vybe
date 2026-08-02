# vybe-test: python/pathlib_runtime/pathlib_hardlink
# origin: languages/python/tests/python/test_pathlib_runtime.rs
# vybe-test-mode: compile

from pathlib import Path
hasattr(Path('.'), 'hardlink_to')
